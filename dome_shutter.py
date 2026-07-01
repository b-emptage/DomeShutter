 # -*- coding: utf-8 -*-
"""
Created on Mon May 18 09:39:38 2026

@author: bemptage
"""
import time
from datetime import datetime, timedelta
import socket
import select
import subprocess
import ctypes
import win32com.client
import pythoncom
#from pydub import AudioSegment
from pydub.playback import play
from queue import Queue
import threading
import sys
import os

TIMEOUT_MINUTES = 5
CHECK_INTERVAL = 60  # seconds between RDP activity checks

last_active_time = datetime.now()

class _Velleman:
    def __init__(self):
        #Load 64 or 32 bit velleman library: 
        if ctypes.sizeof(ctypes.c_voidp)*8 == 64 :  #64bit
            self.lib = ctypes.WinDLL("./K8055fpc64.dll")
        else:  #32bit
            self.lib = ctypes.WinDLL("./K8055D.dll")
        vStatus = self.lib.OpenDevice(0)
        if vStatus <0 :
            raise RuntimeError("Failed to connect to Velleman interface.")
    
    def set_output(self, channel):
        self.lib.SetDigitalChannel(channel)
    
    def clear_output(self, channel):
        self.lib.ClearDigitalChannel(channel)
    
    def read_inputs(self):
        return self.lib.ReadAllDigital()
    
    def read_analogue(self, channel):
        return self.lib.ReadAnalogChannel(channel)

class Dome_Control:
    def __init__(self):
        self.velleman = _Velleman()
        self.e_state = 'stopped'
        self.w_state = 'stopped'
        # output ports for opening/closing
        self.eosw = 6  # east open channel
        self.ecsw = 7  # east close channel
        self.wosw = 4  # west open channel
        self.wcsw = 5  # west close channel
        # input ports for limit switches
        self.wlim = 1  # west shutter limit switch: 0 on, 1 off
        self.elim = 2  # east shutter limit switch: 0 on, 1 off
        self.reed = 3  # dome close reed switch: 1 on, 0 off
        # analog channels for dome position
        self.wpos = 1
        self.epos = 2
        # timers for various timeout functions
        self.e_timer = 0
        self.w_timer = 0
        # direct position values and tolerance for positioning
        self.east_target = None
        self.west_target = None
        self.tolerance = 2
        self.east_position, self.west_position = self._read_shutter_positions()
        
        self.dir_delay = 0.750  # s delay between switching directions
        self.stop_e()  # ensure dome stopped on initialisation
        self.stop_w()
        self._monitor(0.1)  # 100ms polling

    def _set_state(self, shutter, state):
        """
        :param shutter: which shutter to set state for
        :param state: state of shutter
        :return: nothing

        Should be called on any motor event, set the various states and timers appropriately.
        """
        now = time.time()
        if shutter.lower() == 'e':
            self.e_state = state
            self.e_timer = now
            self.last_east = self.east_position
        elif shutter.lower() == 'w':
            self.w_state = state
            self.w_timer = now
            self.last_west = self.west_position

    def open_e(self):
        # stop the motor if the shutter is in the opposite state
        if self.e_state == 'closing':
            self.stop_e()
            time.sleep(self.dir_delay)
        self._set_state('e', 'opening')
        self.velleman.set_output(self.eosw)
        
    def close_e(self):
        if self.e_state == 'opening':
            self.stop_e()
            time.sleep(self.dir_delay)  # pause before swapping directions
        self._set_state('e', 'closing')
        self.velleman.set_output(self.ecsw)

    def stop_e(self):
        self.velleman.clear_output(self.eosw)
        self.velleman.clear_output(self.ecsw)
        self._set_state('e', 'stopped')

    def goto_e(self, setpoint):
        """
        Moves the shutter to the setpoint position
        """
        self.east_target = setpoint
        if self.east_position < self.east_target - self.tolerance:
            # open the dome to increase position. Do nothing if close to target (within tolerance value)
            self.open_e()
        elif self.east_position > self.east_target + self.tolerance:
            self.close_e()
        else:
            # do nothing if setpoint results in too small of a movement
            return
    
    def open_w(self):
        # stop the motor if the shutter is in the opposite state
        if self.w_state == 'closing':
            self.stop_w()
            time.sleep(self.dir_delay)
        self._set_state('w', 'opening')
        self.velleman.set_output(self.wosw)
        
    def close_w(self):
        if self.w_state == 'opening':
            self.stop_w()
            time.sleep(self.dir_delay)
        self._set_state('w', 'closing')
        self.velleman.set_output(self.wcsw)

    def stop_w(self):
        self.velleman.clear_output(self.wosw)
        self.velleman.clear_output(self.wcsw)
        self._set_state('w', 'stopped')

    def goto_w(self, setpoint):
        """
            Moves the west shutter to the setpoint position
        """
        self.west_target = setpoint
        if self.west_position < self.west_target - self.tolerance:
            self.open_w()
        elif self.west_position > self.west_target + self.tolerance:
            self.close_w()
        else:
            return
    
    def _read_shutter_switches(self):
        """
        Checks the status of the dome input switches:
            1: West shutter: 0 on fully open
            2: East shutter: 0 on fully open
            3: Both segments fully closed
        output is sum of the bits
        State values:
            0: Both east and west fully open
            1: East fully open, but west not fully open
            2: West fully open, but east not fully open
            3: East and west not fully open
            7: East and west fully closed
        """
        
        def bit_set(v, chan):
            '''
            Check if a particular bit is set
            '''
            return (v & (1 << (chan - 1))) != 0
            
        val = self.velleman.read_inputs()
        
        return {'west_limit': not bit_set(val, self.wlim),
                'east_limit': not bit_set(val, self.elim),
                'all_closed': bit_set(val, self.reed),
                'raw': val
            }
    
    def _read_shutter_positions(self):
        """
        Returns the current values of the analog inputs on the Velleman
        """
        west = self.velleman.read_analogue(self.wpos)
        east = self.velleman.read_analogue(self.epos)
        return [east,west]
    
    def _update_status(self):
        """
        Reads out the analogue channels of the Velleman board, and checks their
        status. Includes logic to time out and close the digital channels while
        opening or closing
        """
        switches = self._read_shutter_switches()
        self.east_position, self.west_position = self._read_shutter_positions()
        now = time.time()

        # just in case
        if not hasattr(self, 'last_east'):
            self.last_east = self.east_position
        if not hasattr(self, 'last_west'):
            self.last_west = self.west_position

        def _east_set_reached():
            # logic for reaching set position, if one is active on motor activation
            if self.east_target is None:
                # don't want to do anything if no target set
                return False
            if self.e_state == 'opening':
                # When position exceeds set target, we want this to evaluate as true (open past point)
                return self.east_position >= self.east_target - self.tolerance
            elif self.e_state == 'closing':
                # When position smaller than set target, we want to be able to stop motors
                return self.east_position <= self.east_target + self.tolerance
            return False

        def _west_set_reached():
            if self.west_target is None:
                return False
            if self.w_state == 'opening':
                return self.west_position >= self.west_target - self.tolerance
            elif self.w_state == 'closing':
                return self.west_position <= self.west_target + self.tolerance
            return False
        
        # logic for stops on east shutter opening
        if self.e_state == 'opening':
            # update the timer on positive position change
            if self.east_position > self.last_east:
                self.e_timer = now
            # stop shutter if the position hasn't updated for more than 1 sec
            if (now - self.e_timer) > 1:
                print("East shutter opening timeout, stopping")
                self.stop_e()
                self.east_target = None
            # stop if the dome is on the soft limit switch
            elif switches['east_limit']:
                print("East shutter fully open, stopping")
                self.stop_e()
                self.east_target = None
            # stop the motors if a drive to position has been issued and we are close to that position
            elif _east_set_reached():
                self.stop_e()
                self.east_target = None
            # when opening, we only want the analogue channel to increase
            self.last_east = max(self.last_east, self.east_position)
        # logic for stops on east shutter closing
        elif self.e_state == 'closing':
            # update timer on negative position change
            if self.east_position < self.last_east:
                self.e_timer = now
            if (now - self.e_timer) > 1:
                print("East shutter closing timeout, stopping")
                self.stop_e()
                self.east_target = None
            elif switches['all_closed']:
                print("East shutter fully closed, stopping")
                self.stop_e()
                self.east_target = None
            elif _east_set_reached():
                self.stop_e()
                self.east_target = None
            # when closing we only want the analog channel to decrease
            self.last_east = min(self.last_east, self.east_position)
        
        # logic for west shutters, as above
        if self.w_state == 'opening':
            if self.west_position > self.last_west:
                self.w_timer = now
            if (now - self.w_timer) > 1:
                print("West shutter opening timeout, stopping")
                self.stop_w()
                self.west_target = None
            elif switches['west_limit']:
                print("West shutter fully open, stopping")
                self.stop_w()
                self.west_target = None
            elif _west_set_reached():
                self.stop_w()
                self.west_target = None
            self.last_west = max(self.last_west, self.west_position)
        elif self.w_state == 'closing':
            if self.west_position < self.last_west:
                self.w_timer = now
            if (now - self.w_timer) > 1:
                print("West shutter closing timeout, stopping")
                self.stop_w()
                self.west_target = None
            elif switches['all_closed']:
                print("West shutter fully closed, stopping")
                self.stop_w()
                self.west_target = None
            elif _west_set_reached():
                self.stop_w()
                self.west_target = None
            self.last_west = min(self.last_west, self.west_position)

    
    def _monitor(self, pollrate):
        # polling loop to return current dome status
        def loop():
            while True:
                self._update_status()
                time.sleep(pollrate)
        threading.Thread(target=loop, daemon=True).start()
    # TODO set up safe shutdown of thread on close

class Telescope:
    ASCOM_DRIVER = "DDR_ASCOM.Telescope"
    def __init__(self):
        self.tele = None
        try:
            pythoncom.CoInitialize()  # Initialize COM
            self.tele = win32com.client.Dispatch(self.ASCOM_DRIVER)
        except Exception as e:
            print(f"Error connecting to ASCOM driver {self.ASCOM_DRIVER}: {e}")

    def connect(self):
        if self.tele:
            self.tele.Connected = True
    
    def check_park(self):
        if self.tele and self.tele.Connected:
            return self.tele.AtPark
        
    def park(self):
        if self.tele and self.tele.Connected:
            # only park telescope if not currently parked
            if not self.check_park():
                self.tele.Park()
        
    def disconnect(self):
        if self.tele:
            self.tele.Connected = False        
            
    def abort_slew(self):
        # only abort slew if telescope already slewing
        if self.tele and self.check_slew_status():
            self.tele.AbortSlew()
    
    def check_slew_status(self):
        if self.tele and self.tele.Connected:
            return self.tele.Slewing

class TCPServer:
    HOST = '192.168.1.1'
    PORT = 1338
    
    _instance = None  # Class variable to store the single instance

    def __new__(cls, host=HOST, port=PORT):
        # Ensure that only one instance is created
        if cls._instance is None:
            cls._instance = super(TCPServer, cls).__new__(cls)
            cls._instance.__init__(host, port)  # Initialize the server only once
        return cls._instance
    
    def __init__(self, host=HOST, port=PORT):
        # The constructor will only be called the first time
        if not hasattr(self, 'server_socket'):  # Check if initialized already
            self.server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.server_socket.bind((host, port))
            self.server_socket.listen(1)
            print(f"Server listening on {host}:{port}")
            # SETBLOCKING LETS THE REST OF THE PROGRAM RUN WITHOUT WAITING
            self.server_socket.setblocking(False)
            self.client_socket = None
            self.client_addr = None
            self.polling = False

    def start_polling(self, root, callback):
        if not self.polling:
            self.polling = True
            self._poll_server(root, callback)

    def _poll_server(self, root, callback):
        self.accept_client()
        data = self.receive_data()
        if data:
            callback(data)
        elif not self.client_socket:  # Client is disconnected
            callback(None)  # Notify of disconnection
        root.after(500, self._poll_server, root, callback)
    
    def reinitialize_server(self):
        print("Reinitializing the server socket...")
        self.close()
        self.server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.server_socket.bind((self.HOST, self.PORT))
        self.server_socket.listen(1)
        self.server_socket.setblocking(False)

    def accept_client(self):
        if self.server_socket is None or self.server_socket.fileno() == -1:
            print("Server socket is invalid. Reinitializing...")
            self.reinitialize_server()
            return
        # Accept client connection if available
        try:
            self.client_socket, self.client_addr = self.server_socket.accept()
            self.client_socket.setblocking(False)  # Make client socket non-blocking
            print(f"Connected by {self.client_addr}")
            self.client_socket.sendall(b"Successfully connected to Dome Shutter Controller")
        except BlockingIOError:
            pass  # No connection ready to be accepted

    def receive_data(self):
        """Receive data from the client if available."""
        if self.client_socket:
            try:
                # Check if client socket is valid before using it
                if self.client_socket.fileno() == -1:
                    print("Client socket is invalid. Cleaning up...")
                    self.cleanup_client()
                    return None
                
                # Check if data is available to read
                readable, _, _ = select.select([self.client_socket], [], [], 0)
                if readable:
                    data = self.client_socket.recv(1024)
                    if data:                        
                        # Process specific commands
                        if b'OK' in data:
                            self.client_socket.send(b'Host: OK')
                        
                        return data.decode()
                    else:
                        # No data means the client has disconnected
                        print(f"{self.client_addr} disconnected.")
                        self.cleanup_client()
                return None
    
            except BlockingIOError:
                pass  # No data ready to be read
            except ConnectionResetError:
                print("Connection forcibly closed by client. Cleaning up...")
                self.cleanup_client()
            except Exception as e:
                print(f"Unexpected error while receiving data: {e}")
                self.cleanup_client()
        return None
    
    def cleanup_client(self):
        # Clean up client resources after disconnection or error.
        if self.client_socket:
            try:
                self.client_socket.close()
            except Exception as e:
                print(f"Error closing client socket: {e}")
            finally:
                self.client_socket = None
                self.client_addr = None
        print("Client resources cleaned up. Ready to accept a new connection.")
    
    def send_data(self, message):
        # send some data to the client if the socket is writable
        if self.client_socket:
            try:
                _, writable, _ = select.select([], [self.client_socket], [], 1)
                if writable:
                    self.client_socket.send(message.encode('utf-8'))
            except Exception as e:
                print(f"Error sending data to {self.client_socket}: {e}")
        return None

    def close(self):
        # close the client and host sockets
        if self.client_socket:
            try:
                self.client_socket.close()
            except Exception as e:
                print(f"Error closing client socket: {e}")
            finally:
                self.client_socket = None
                self.client_addr = None
        if self.server_socket:
            try:
                self.server_socket.close()
            except Exception as e:
                print(f"Error closing server socket: {e}")
            finally:
                self.server_socket = None
        print("Sockets closed.")


class Speaker:
    def __init__(self):
        self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
        self.speaker.Rate = -3
        self.queue = Queue()
        self.thread = threading.Thread(target=self._process_queue, daemon=True)
        self.thread.start()

    def _process_queue(self):
        pythoncom.CoInitialize()
        while True:
            item = self.queue.get()
            if item is None:  # Stop signal
                break
    
            file_path = None  # Ensure it's always defined
            try:
                if isinstance(item, str):  # Text input
                    if self.speaker:
                        self.speaker.Speak(item)
                    else:
                        print(f"[{time.strftime('%H:%M:%S')}] "+
                              "Speaker not initialized. Skipping speech.")
                elif isinstance(item, dict) and item.get("type") == "audio":  # Audio input
                    file_path = item.get("file_path")
                    if file_path:
                        self._play_audio(file_path)
                    else:
                        print("No audio file provided, skipping play.")
                else:
                    print(f"Invalid input to Speaker: {item}")
            except Exception as e:
                print(f"[{time.strftime('%H:%M:%S')}] "+
                      f"Error processing item: {e}")
        pythoncom.CoUninitialize()

    def _play_audio(self, file_path):
        try:
            audio = AudioSegment.from_file(file_path)
            play(audio)
        except Exception as e:
            print(f"[{time.strftime('%H:%M:%S')}] "
                  + f"Error playing audio file {file_path}: {e}")

    def speak_async(self, text):
        #queue a text message for text-to-speech.
        self.queue.put(text)

    def play_audio_async(self, file_path):
        # queue an audio file to play.
        self.queue.put({"type": "audio", "file_path": file_path})

    def shutdown(self):
        #Shutdown the speaker system
        self.queue.put(None)
        self.thread.join()

class shutdown_monitor:
    def __init__(self):
        if os.path.exists("stop_monitor.flag"):
            os.remove("stop_monitor.flag")
        self.running = True
        global last_active_time 
        last_active_time = datetime.now()

    def is_rdp_active(self):
        '''
        Checks if there is currently an active RDP session to the machine
        '''
        try:
            result = subprocess.run(['qwinsta'], capture_output=True, text=True)
            output = result.stdout.lower()
            lines = output.splitlines()
            for line in lines:
                if 'rdp-tcp' in line and 'active' in line:
                    return True
            return False
        except Exception as e:
            print(f"Error checking RDP session: {e}")
            return False

    def disconnected_client_shutdown(self):
        print(f"[{datetime.now()}] No RDP session detected for {TIMEOUT_MINUTES}m.")
        scope = None
        dome = None
        try:
            # initialise a connection to the dome, telescope
            dome = Dome_Control()
            scope = Telescope()
        except Exception as e:
            print(f"Error connecting equipment: {e}")
        if scope:
            scope.abort_slew()
            time.sleep(0.5)  # give it a half second
            scope.park()
        if dome:
            # close west first, as telescope sits on west side in park state
            dome.close_w()
            time.sleep(1)  # prevent over-current on switch by staggering motors
            dome.close_e()
        pythoncom.CoUninitialize()  # this will close the ASCOM client

    def monitor_rdp(self):
        global last_active_time
        print("Starting RDP session monitor...")
        flag_changed = False  # to restart timer when the stop flag is removed
        # as long as stop_monitor.flag file doesn't exist, monitor RDP sessions
        while self.running:
            stop_flag = os.path.exists("stop_monitor.flag")
            # if stop_flag is true we don't want to do anything
            if not stop_flag:
                # restart the timer if stop_flag is True (set in previous loop)
                if flag_changed:
                    last_active_time = datetime.now()
                if self.is_rdp_active():
                    # we have an active connection, update last active time
                    last_active_time = datetime.now()
                    print(f"[{datetime.now()}] RDP session active.")
                else:
                    # calculate time between last active time and now
                    elapsed = datetime.now() - last_active_time
                    print(f"[{datetime.now()}] No active RDP session. Elapsed: {elapsed}")
                    # check if RDP has been inactive for more than set timeout
                    if elapsed > timedelta(minutes=TIMEOUT_MINUTES):
                        self.disconnected_client_shutdown()
                        # Wait until RDP reconnects
                        self.wait_for_rdp()
            # stop_flag exists, which means something wants the monitor stopped
            else:
                with open("stop_monitor.flag", "r") as f:
                    contents = f.read().strip().lower()
                    if "end" in contents:
                        print("End flag present from main loop. Stopping RDPM")
                        self.running = False
                        break
            # this will turn true when stop_flag becomes true (was changed) 
            flag_changed = stop_flag
            time.sleep(CHECK_INTERVAL)
    
    def wait_for_rdp(self):
        print("Waiting for RDP reconnection...")
        while not self.is_rdp_active() and self.running:
            print(f"[{datetime.now()}] No active RDP session.")
            # check for the stop flag and stop if it exists
            if os.path.exists("stop_monitor.flag"):
                with open("stop_monitor.flag", "r") as f:
                    contents = f.read().strip().lower()
                    if "end" in contents:
                        print("End flag present. Stopping RDP monitor")
                        self.running = False
                        break
            time.sleep(CHECK_INTERVAL)
        # restart monitor when session is connected
        if self.is_rdp_active() and self.running:
            print(f"[{datetime.now()}] RDP session reconnected.")
            return  # go back to monitor_rdp loop


def start_rdp_monitor():
    monitor = shutdown_monitor()
    threading.Thread(target=monitor.monitor_rdp, daemon=True).start()

shutdown_mode = {"enabled": False}
admin_window_instance = None

#TODO: move to QT architecture and into dome_UI
def open_admin_menu():
    global admin_window_instance
    
    if admin_window_instance is not None and \
        tk.Toplevel.winfo_exists(admin_window_instance):
        admin_window_instance.lift()
        return
    
    admin_window_instance = tk.Toplevel()
    admin_window_instance.title("Shutdown Controls")
    admin_window_instance.text = tk.LabelFrame(admin_window_instance)
    admin_window_instance.geometry("200x150")
    
    def on_close():
        # make sure single instance performs correctly on close
        global admin_window_instance
        admin_window_instance.destroy()
        admin_window_instance = None
    
    admin_window_instance.protocol("WM_DELETE_WINDOW", on_close)

    def toggle_sd_mode():
        shutdown_mode["enabled"] = sd_var.get()
        print(f"Toggle set to: {shutdown_mode['enabled']}")
        # turn off RDP monitor
        if shutdown_mode["enabled"]:
            with open("stop_monitor.flag", "w") as f:
                f.write(f"{datetime.now()}: stop")
                z = os.path.exists("stop_monitor.flag")
                print(f'Creating stop flag: {z}')
        elif not shutdown_mode["enabled"]:
            if os.path.exists("stop_monitor.flag"):
                os.remove("stop_monitor.flag")
    # make sure people know what ticking the button will do
    warn_label = tk.Label(admin_window_instance,
                          text="WARNING: Disabling this feature\n" +
                          "will stop the system from \n" +
                          "shutting down safely in \n" +
                          "the event of a disconnect.")
    warn_label.pack(pady=10)    

    sd_var = tk.BooleanVar(value=shutdown_mode["enabled"])
    sd_checkbox = ttk.Checkbutton(admin_window_instance,
                                   text="Disable RDP auto-shutdown?",
                                   variable=sd_var, command=toggle_sd_mode)
    sd_checkbox.pack(pady=10)