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
#from msl.loadlib import LoadLibrary
#from msl.loadlib import IS_PYTHON_64BIT
#import os
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
        if self.client_socket:
            try:
                # Check if client socket is valid before select
                if self.client_socket.fileno() == -1:
                    print("Client socket is invalid. Cleaning up...")
                    self.client_socket.close()
                    self.client_socket = None
                    self.client_addr = None
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
                        # No data indicates the client has disconnected
                        print(f"{self.client_addr} disconnected")
                        self.client_socket.close()
                        self.client_socket = None
                        self.client_addr = None
            except BlockingIOError:
                pass  # No data ready to be read
            except ConnectionAbortedError:
                print("Connection aborted by client. Reinitializing server...")
                self.reinitialize_server()
        return None
    
    def send_data(self, message):
        if self.client_socket:
            try:
                _, writable, _ = select.select([], [self.client_socket], [], 0.5)
                if writable:
                    self.client_socket.send(message.encode('utf-8'))
            except Exception as e:
                print(f"Error sending data to {self.client_socket}: {e}")
        return None

    def close(self):
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
    oTime = 9.0 #Time for shutter to fully open or Close
    timeInc=int(1000*oTime/100.0) #time between callbacks for shutter progress update
    timeCorrection = 0.11 #Seconds to add to opening/closing time for each stop
    shutterColumn = 0    #incremented for each Shutter instance
    tcp_polling = False
    
    def __init__(self,root,titleText,side):
        self.root = root
        self.name = titleText        
        self.tStart = time.time()
        self.tEnd = Shutter.oTime
        self.cb = "None"  #Holds callback ID for shutter timing
        self.upDown = -1 #initially closed
        
        self.tcp_server = TCPServer()
        # ensure we only establish a single poller
        if not Shutter.tcp_polling:
            self.poll_server()
            Shutter.tcp_polling = True
        
        self.lf = ttk.LabelFrame(root,labelanchor='n',text=titleText,style="TLabelframe")
        self.lf.grid(column=Shutter.shutterColumn,row=0,padx=1,pady=1)
        Shutter.shutterColumn += 1
        self.pb = ttk.Progressbar(
            self.lf,
            orient='vertical',
            mode='determinate',
            length=100,
            value=100,
            style="red.Horizontal.TProgressbar"
        )
        # place the progressbar
        self.value_label = ttk.Label(self.lf, text=self.update_progress_label(),width=14)
        if side == "Left":
            self.pb.grid(column=0, row=0, columnspan=1,rowspan=3, padx=10, pady=10)
            self.value_label.grid(column=0, row=3, columnspan=1)
        else:
            self.pb.grid(column=2, row=0, columnspan=1,rowspan=3, padx=10, pady=10)
            self.value_label.grid(column=2, row=3, columnspan=1)
        
        # start button
        self.open_button = ttk.Button(self.lf,text='Open',command=self.openShutter)
        self.open_button.grid(column=1, row=2, padx=10, pady=2, sticky=tk.E)
        
        
        self.close_button = ttk.Button(self.lf,text='Close',command=self.closeShutter)
        self.close_button.grid(column=1, row=0, padx=10, pady=2, sticky=tk.E)
        
        self.stop_button = ttk.Button(self.lf,text='Stop',command=self.stopShutter)
        self.stop_button.grid(column=1, row=1, padx=10, pady=2, sticky=tk.W)
        
        if self.name == "West":
            self.motor = Motor(Motor.WESTOPEN,Motor.WESTCLOSE)
        elif self.name == "East":
            self.motor = Motor(Motor.EASTOPEN,Motor.EASTCLOSE)
 
    def update_progress_label(self):
        return "Closed: {:5.1f}%".format(self.pb['value'])
    
    def poll_server(self):
        # Check for incoming connections and data
        self.tcp_server.accept_client()
        data = self.tcp_server.receive_data()
        
        if data:
            if "close" in data.lower():
                # need to close both shutters
                westShutter.closeShutter()
                # stagger shutter closes to prevent overcurrent on system
                root.after(1500, eastShutter.closeShutter)
                self.tcp_server.send_data("Dome closing")
    
        # Schedule the next poll in 500 ms
        self.root.after(500, self.poll_server)
    
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