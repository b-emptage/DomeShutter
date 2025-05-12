# -*- coding: utf-8 -*-
"""
Created on Tue Oct 26 03:10:51 2021

@author: Kym
"""

from tkinter import ttk
import tkinter as tk
from tkinter.messagebox import askokcancel
import time
import socket
import select
from ctypes import *
import win32com.client
from pydub import AudioSegment
from pydub.playback import play
from queue import Queue
import threading
#from msl.loadlib import LoadLibrary
#from msl.loadlib import IS_PYTHON_64BIT
import sys


class Motor():
    #Velleman output ports
    WESTOPEN = 4
    WESTCLOSE = 5
    EASTOPEN = 6
    EASTCLOSE = 7
    #Delays
    dirDelay = 750  #msec delay when changing motor direction
    offDelay = 1500 #Time to leave power on motor once opened/closed
    
#    velleman=windll.LoadLibrary (os.getcwd()+"\\K8055D")
#    Note that WinDLL is the same as windll.LoadLibrary
    #Load 64 or 32 bit velleman library: 
    if sizeof(c_voidp)*8 == 64 :  #64bit
        velleman=WinDLL("./K8055fpc64.dll")
    else:  #32bit
        velleman=WinDLL("./K8055D.dll")
    #velleman=LoadLibrary("K8055D.dll","windll").lib
    vStatus = velleman.OpenDevice(0)
    if vStatus <0 :
        master = tk.Tk()
        master.withdraw()       
        if not askokcancel("DomeShutter info.",message=
                  "Failed to connect to Velleman interface.\n"+
                  "Shutter control wont work without the connection\n"+
                  "Hit Ok to continue, or Cancel?"):
            sys.exit()
            #raise SystemExit(0)
        master.destroy()   
    def __init__(self, openSW,closeSW):
        self.openSW = openSW
        self.closeSW = closeSW
        self.stop()
        self.closing = True
        self.opening = True

    def open(self):
        self.opening = True
        Motor.velleman.SetDigitalChannel(self.openSW)
        
    def close(self):
        self.closing = True
        Motor.velleman.SetDigitalChannel(self.closeSW)

    def stop(self):
        Motor.velleman.ClearDigitalChannel(self.openSW)
        Motor.velleman.ClearDigitalChannel(self.closeSW)
        self.closing = False
        self.opening = False

 
#velleman=windll.LoadLibrary (os.getcwd()+"\\K8055D")
#velleman.OpenDevice(0)
#mydll.ConfigureMCP2200(0,9600,0,0,False,False, False, False)

class TCPServer():
    HOST = '127.0.0.1'
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
                        print(f"Client: {data.decode()}")
                        
                        # Process specific commands
                        if data == b'beep':
                            self.client_socket.send(b'boop')
                        
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

class Shutter():
    oTime = 9.0  # Time for shutter to fully open or close
    timeInc = int(1000 * oTime / 100.0)  # Time between callbacks for shutter progress update
    timeCorrection = 0.11  # Seconds to add to opening/closing time for each stop
    shutterColumn = 0  # Incremented for each Shutter instance
    tcp_indicator_frame = None
    tcp_status_label = None
    tcp_led_canvas = None
    
    def __init__(self, root, titleText, side):
        self.root = root
        self.name = titleText
        self.tStart = time.time()
        self.tEnd = Shutter.oTime
        self.cb = "None"  # Holds callback ID for shutter timing
        self.upDown = -1  # Initially closed
        
        self.tcp_server = TCPServer()
        self.tcp_server.start_polling(self.root, self.handle_tcp_data)
        
        # Create a frame to hold everything for this shutter
        self.outer_frame = tk.Frame(root, bg="#990000")
        self.outer_frame.grid(column=Shutter.shutterColumn, row=0, padx=2, pady=2, sticky="nsew")

        # Add the main label frame for the shutter
        self.lf = ttk.LabelFrame(self.outer_frame, labelanchor="n", text=titleText, style="TLabelframe")
        self.lf.pack(fill="both", expand=True, padx=2, pady=2)
        Shutter.shutterColumn += 1
        
        self.pb = ttk.Progressbar(
            self.lf,
            orient='vertical',
            mode='determinate',
            length=100,
            value=100,
            style="red.Horizontal.TProgressbar"
        )
        # Place the progress bar
        self.value_label = ttk.Label(self.lf, text=self.update_progress_label(), width=14)
        if side == "Left":
            self.pb.grid(column=0, row=0, columnspan=1, rowspan=3, padx=10, pady=10)
            self.value_label.grid(column=0, row=3, columnspan=1)
        else:
            self.pb.grid(column=2, row=0, columnspan=1, rowspan=3, padx=10, pady=10)
            self.value_label.grid(column=2, row=3, columnspan=1)
        
        # Start button
        self.open_button = ttk.Button(self.lf, text='Open', command=self.openShutter)
        self.open_button.grid(column=1, row=2, padx=10, pady=2, sticky=tk.E)
        
        self.close_button = ttk.Button(self.lf, text='Close', command=self.closeShutter)
        self.close_button.grid(column=1, row=0, padx=10, pady=2, sticky=tk.E)
        
        self.stop_button = ttk.Button(self.lf, text='Stop', command=self.stopShutter)
        self.stop_button.grid(column=1, row=1, padx=10, pady=2, sticky=tk.W)
        
        if self.name == "West":
            self.motor = Motor(Motor.WESTOPEN, Motor.WESTCLOSE)
        elif self.name == "East":
            self.motor = Motor(Motor.EASTOPEN, Motor.EASTCLOSE)

    def update_progress_label(self):
        return "Closed: {:5.1f}%".format(self.pb['value'])
    
    def handle_tcp_data(self, data):
        # Handle data from the TCP server
        if data:
            if "close" in data.lower():
                # Close both shutters
                speaker.speak_async("Rain detected. Closing dome in 5 seconds - please stand clear.")
                self.root.after(5000, westShutter.closeShutter)
                self.root.after(7000, eastShutter.closeShutter)
                self.tcp_server.send_data("Dome closing")
            self.update_tcp_status(True)
        else:
            self.update_tcp_status(False)

    def update_tcp_status(self, connected):
        """Update the TCP LED indicator."""
        color = "#009900" if connected else "#990000"
        # Can have an issue on startup, where east shutter isn't initialised yet
        try:
            eastShutter.outer_frame.config(bg=color)
        except Exception:
            pass  # just continue on if we can't set the background colour
        try:
            westShutter.outer_frame.config(bg=color)
        except Exception:
            pass  # keep on keeping on
        # Propagate color changes to internal widgets if needed
        # self.lf.config(style="Custom.TLabelframe")
        # ttk.Style().configure("Custom.TLabelframe", background=color)
        # ttk.Style().configure("Custom.TLabelframe.Label", background=color)

    
    def on_close(self):
        # Close the server and the GUI
        self.tcp_server.close()
    
    def status(self):
        if self.cb in root.tk.call('after','info'):
            root.after_cancel(self.cb)
        #Are we opening or closing
        if (self.upDown==-1 and self.pb['value'] >0) or (self.upDown==1 and self.pb['value']<100):
            percent = (self.tEnd - (time.time()-self.tStart))/Shutter.oTime
            #print "Yeh",self.pb['value'],self.upDown,percent
            if self.upDown==-1:
                percent = 100*percent
            else:
                percent = 100*(1-percent)
            if percent<=0:
                percent=0.0
            elif percent>=100.0:
                #self.open_button.configure(style="TButton")
                percent = 100.0
            self.pb['value'] = percent
            #pbE['value'] += upDown
            self.value_label['text'] = self.update_progress_label()
            self.cb = root.after(Shutter.timeInc,self.status)
        elif self.pb['value'] <= 0:
            #self.open_button.configure(style="TButton")
            self.open_button.configure(style="green.TButton")
            #showinfo(message='East OPEN')
            #self.upDown = 1
            self.cb = root.after(Motor.offDelay,self.motor.stop)
        elif self.pb['value'] >= 100:
            #self.close_button.configure(style="TButton")
            self.close_button.configure(style="green.TButton")
            #showinfo(message='East CLOSED')
            #self.upDown = -1
            self.cb = root.after(Motor.offDelay,self.motor.stop)
        #print "status",self.pb['value']


    def openShutter(self):
        if self.cb in root.tk.call('after','info'):
            root.after_cancel(self.cb)
        if self.motor.closing:
            self.stopShutter()
            self.cb=root.after(Motor.dirDelay,self.openShutter)
            return
        if not self.motor.opening:
            self.motor.open()
            percent = self.pb['value']
            if (percent <= 0.0) or (percent >=100.0):
                delta=0.0
            else:
                delta=Shutter.timeCorrection
            self.tStart = time.time()+delta
            self.tEnd=Shutter.oTime*percent/100.0
            self.upDown=-1
        self.cb = root.after(Shutter.timeInc,self.status)
        self.tcp_server.send_data("Dome opened")
        self.open_button.configure(style="red.TButton")
        self.close_button.configure(style="TButton")

    def closeShutter(self):
        if self.cb in root.tk.call('after','info'):
            root.after_cancel(self.cb)
        if self.motor.opening:
            self.stopShutter()
            self.cb = root.after(Motor.dirDelay,self.closeShutter)
            return
        if not self.motor.closing:
            self.motor.close()
            percent = self.pb['value']
            if (percent <= 0.0) or (percent >=100.0):
                delta=0.0
            else:
                delta=Shutter.timeCorrection
            self.tStart = time.time()+delta
            self.tEnd=Shutter.oTime*(100-percent)/100.0
            self.upDown=1
        self.cb = root.after(Shutter.timeInc,self.status)
        self.close_button.configure(style="red.TButton")
        self.open_button.configure(style="TButton")
        #print "close",self.upDown,self.tEnd
        
    def stopShutter(self):
#        if self.pb['value']<=0:
#            self.pb.stop()
        root.after_cancel(self.cb)
        self.motor.stop()
        self.value_label['text'] = self.update_progress_label()
        if self.pb['value'] >= 100.0:
            self.close_button.configure(style="green.TButton")
            self.open_button.configure(style="TButton")
        elif self.pb['value'] <= 0.0:
            self.open_button.configure(style="green.TButton")
            self.close_button.configure(style="TButton")
        else:
            self.open_button.configure(style="TButton")
            self.close_button.configure(style="TButton")

# TODO: fix speaking (duplicate to RainMonT)
class Speaker:
    def __init__(self):
        self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
        self.speaker.Rate = -3
        self.queue = Queue()
        self.thread = threading.Thread(target=self._process_queue, daemon=True)
        self.thread.start()

    def _process_queue(self):
        while True:
            item = self.queue.get()
            if item is None:  # Stop signal
                break
            try:
                if isinstance(item, str):  # Text input
                    self.speaker.Speak(item)
                elif isinstance(item, dict) and item.get("type") == "audio":  # Audio input
                    file_path = item.get("file_path")
                    self._play_audio(file_path)
                else:
                    print(f"Invalid input to Speaker: {item}")
            except Exception as e:
                print(f"Error processing queue: {e}")

    def _play_audio(self, file_path):
        try:
            audio = AudioSegment.from_file(file_path)
            play(audio)
        except Exception as e:
            print(f"Error playing audio file {file_path}: {e}")

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
        


# root window
root = tk.Tk()
#root.geometry('370x175')
root.title('Shutter Control v0.9.2')
# progressbar
#s = ttk.Style()
#print s.theme_names()
ttk.Style().theme_use('winnative')
#s.theme_use('xpnative')
#s.theme_use('vista')
pbc="#666666"
ttk.Style().configure("red.Horizontal.TProgressbar", foreground=pbc, background=pbc,troughcolor ='Darkblue')

#bColour = ttk.Style().lookup('openE_button', 'background')
#print bColour
ttk.Style().configure("red.TButton", background="red",foreground="#cc3333")
ttk.Style().configure("green.TButton", background="green")
#ttk.Style().map("red.TButton",
#    foreground=[('pressed', 'red'), ('active', 'blue')],
#    background=[('pressed', '!disabled', 'black'), ('active', 'red')]
#    )

#Background styles darker than default and LableFrame text to bold
bg = "#aaaaaa"
#print ttk.Style().lookup("TLabelframe.Label", "font")
ttk.Style().configure('TLabelframe', background=bg)
ttk.Style().configure('TLabelframe.Label', background=bg,font=("TkDefaultFont", 9, 'bold'))
ttk.Style().configure('TLabel', background=bg)

#Create an instance of each shutter
speaker = Speaker()
westShutter = Shutter(root, "West", "Left")
eastShutter = Shutter(root, "East", "Right")
root.resizable(False, False)

# make sure the tcp server is closing
def on_close():
    if eastShutter:
        eastShutter.on_close()
    if westShutter:
        westShutter.on_close()
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_close)

root.mainloop()

Motor.velleman.CloseDevice()