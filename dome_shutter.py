# -*- coding: utf-8 -*-
"""
Created on Mon May 18 09:39:38 2026

@author: bemptage
"""

from tkinter import ttk
import tkinter as tk
from tkinter.messagebox import askokcancel
import time
from datetime import datetime, timedelta
import socket
import select
import subprocess
import ctypes
import win32com.client
import pythoncom
from pydub import AudioSegment
from pydub.playback import play
from queue import Queue
import threading
#from msl.loadlib import LoadLibrary
#from msl.loadlib import IS_PYTHON_64BIT
import sys
import os


class Velleman:
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

class Dome_Control():
    #Delays
    offDelay = 1500 #Time to leave power on motor once opened/closed
    
    def __init__(self):
        self.velleman = Velleman()
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
        # placeholder for position data
        self.last_pos = [0,0]
        
        self.dir_delay = 0.750  # s delay between switching directions
        self.stop_e()  # ensure dome stopped on initialisation
        self.stop_w()
        self.monitor(0.1)  # 100ms polling
        

    def open_e(self):
        # stop themotor if the hutter is in the opposite state
        if self.e_state == 'closing':
            self.stop_e()
            time.sleep(self.dir_delay)
        self.e_state = 'opening'
        self.velleman.set_output(self.eosw)
        
    def close_e(self):
        if self.e_state == 'opening':
            self.stop_e()
            time.sleep(self.dir_delay)  # pause before swapping directions
        self.e_state = 'closing'
        self.velleman.set_output(self.ecsw)

    def stop_e(self):
        self.velleman.clear_output(self.eosw)
        self.velleman.clear_output(self.ecsw)
        self.e_state = 'stopped'
    
    def open_w(self):
        # stop the motor if the shutter is in the opposite state
        if self.w_state == 'closing':
            self.stop_w()
            time.sleep(self.dir_delay)
        self.w_state = 'opening'
        self.velleman.set_output(self.wosw)
        
    def close_w(self):
        if self.w_state == 'opening':
            self.stop_w()
            time.sleep(self.dir_delay)
        self.w_state = 'closing'
        self.velleman.set_output(self.wcsw)

    def stop_w(self):
        self.velleman.clear_output(self.wosw)
        self.velleman.clear_output(self.wcsw)
        self.w_state = 'stopped'
    
    def read_shutter_switches(self):
        '''
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
        '''
        
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
    
    def read_shutter_positions(self):
        '''
        Returns the current values of the analog inputs on the Velleman
        '''
        west = self.velleman.read_analogue(self.wpos)
        east = self.velleman.read_analogue(self.epos)
        return [east,west]
    
    def update_status(self):
        switches = self.read_shutter_switches()
        self.east_position, self.west_position = self.read_shutter_positions()
        now = time.time()
        if not hasattr(self, 'e_timer'):
            self.e_timer = now
        if not hasattr(self, 'w_timer'):
            self.w_timer = now
        if not hasattr(self, 'last_east'):
            self.last_east = self.east_position
        if not hasattr(self, 'last_west'):
            self.last_west = self.west_position
        
        # logic for stops on east shutter opening
        if self.e_state == 'opening':
            # update the timer on positive position change
            if self.east_position > self.last_east:
                self.e_timer = now
            # stop shutter if the position hasn't updated for more than 1 sec
            if (now - self.e_timer) > 1:
                print("East shutter opening timeout, stopping")
                self.stop_e()
            # stop if the dome is on the soft limit switch
            if switches['east_limit']:
                print("East shutter fully open, stopping")
                self.stop_e()
            # when opening, we only want the analogue channel to increase
            self.last_east = max(self.last_east, self.east_position)
        # logic for stops on east shutter closing
        if self.e_state == 'closing':
            # update timer on negative position change
            if self.east_position < self.last_east:
                self.e_timer = now
            if (now - self.e_timer) > 1:
                print("East shutter closing timeout, stopping")
                self.stop_e()
            if switches['all_closed']:
                print("East shutter fully closed, stopping")
                self.stop_e()
            # when closing we only want the analog channel to decrease
            self.last_east = min(self.last_east, self.east_position)
        
        # logic for west shutters, as above
        if self.w_state == 'opening':
            if self.west_position > self.last_west:
                self.w_timer = now
            if (now - self.w_timer) > 1:
                print("West shutter opening timeout, stopping")
                self.stop_w()
            if switches['west_limit']:
                print("West shutter fully open, stopping")
                self.stop_w()
            self.last_west = max(self.last_west, self.west_position)
        if self.w_state == 'closing':
            if self.west_position < self.last_west:
                self.w_timer = now
            if (now - self.w_timer) > 1:
                print("West shutter closing timeout, stopping")
                self.stop_w()
            if switches['all_closed']:
                print("West shutter fully closed, stopping")
                self.stop_w()
            self.last_west = min(self.last_west, self.west_position)

    
    def monitor(self, pollrate):
        # polling loop to return current dome status
        def loop():
            while True:
                self.update_status()
                time.sleep(pollrate)
        threading.Thread(target=loop, daemon=True).start()