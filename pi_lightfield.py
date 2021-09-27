# -*- coding: utf-8 -*-
"""
Created on Tue May 18 15:26:36 2021

@author: Neo
"""
import clr
import sys
import os
from System.IO import *
from System import String
clr.AddReference('System.Collections')
clr.AddReference('System.Threading')
import System
from System.Collections.Generic import List
from System.Runtime.InteropServices import Marshal
from System.Runtime.InteropServices import GCHandle, GCHandleType

from System.Threading import AutoResetEvent


sys.path.append(os.environ['LIGHTFIELD_ROOT'])
sys.path.append(os.environ['LIGHTFIELD_ROOT']+"\\AddInViews")
clr.AddReference('PrincetonInstruments.LightFieldViewV5')
clr.AddReference('PrincetonInstruments.LightField.AutomationV5')
clr.AddReference('PrincetonInstruments.LightFieldAddInSupportServices')


from PrincetonInstruments.LightField.Automation import Automation
from PrincetonInstruments.LightField.AddIns import ExperimentSettings
from PrincetonInstruments.LightField.AddIns import CameraSettings
from PrincetonInstruments.LightField.AddIns import SpectrometerSettings
from PrincetonInstruments.LightField.AddIns import DeviceType
from PrincetonInstruments.LightField.AddIns import ImageDataFormat

import numpy as np
import time
import copy
import xml.dom.minidom


def set_value(setting, value):    
    # Check for existence before setting
    # gain, adc rate, or adc quality
    if experiment.Exists(setting):
        experiment.SetValue(setting, value)

def device_found():
    # Find connected device
    for device in experiment.ExperimentDevices:
        if (device.Type == DeviceType.Camera):
            return True
     
    # If connected device is not a camera inform the user
    print("Princeton Instruments camera not found. Please add a camera and try again.")
    return False


__imdata = [0]
def experimentDataReady(sender, event_args):
    while not experiment.IsReadyToRun and experiment.IsRunning:
        time.sleep(0.1)
    array = np.array(list(event_args.ImageDataSet.GetFrame(0,0).GetData()))
    __imdata[0] = array
        
def experiment_completed(sender, event_args):
    acquireCompleted.Set()
    experiment.ImageDataSetReceived -= experimentDataReady

def experiment_setting_changed(sender, event_args):
    update_height_width()
    

#%%
def get_spectrum():
    if experiment.IsReadyToRun and not experiment.IsRunning:
        experiment.ImageDataSetReceived += experimentDataReady
        experiment.TakeOneLook()
        acquireCompleted.WaitOne()
        array1D = copy.copy(__imdata[0])
        ans = _convert_array1D_to_data(array1D)
        return ans
    else:
        time.sleep(0.1)
        return get_spectrum()

def get_spectrum_direct():
    if experiment.IsReadyToRun and not experiment.IsRunning:
        imdataset = experiment.Capture(1)
        buf = imdataset.GetFrame(0,0).GetData()
        array1D = np.array(list(buf))
        ans = _convert_array1D_to_data(array1D)
        return ans
    else:
        time.sleep(0.1)
        return get_spectrum_direct()

def get_spectrum_autosave():
    if experiment.IsReadyToRun and not experiment.IsRunning:
        experiment.ImageDataSetReceived += experimentDataReady
        experiment.Acquire()
        acquireCompleted.WaitOne()
    else:
        time.sleep(0.1)
        get_spectrum_autosave()
    
def get_wls():
    return np.array(list(experiment.SystemColumnCalibration))

def get_wls_from_recently_saved_spec():
    imdataset = filemanager.OpenFile(filemanager.GetRecentlyAcquiredFileNames()[0],System.IO.FileAccess.Read)
    XML = filemanager.GetXml(imdataset)
    dom = xml.dom.minidom.parseString(XML)
    wls = dom.getElementsByTagName('Wavelength')[0].childNodes[0].data
    wls = np.r_[wls.split(',')].astype(float)
    return wls

def read_spectrum_from_recently_saved_spec():
    imdataset = filemanager.OpenFile(filemanager.GetRecentlyAcquiredFileNames()[0],System.IO.FileAccess.Read)
    array1D = np.array(list(imdataset.GetFrame(0,0).GetData()))
    ans = _convert_array1D_to_data(array1D)
    return ans

def set_exp_time_ms(ms):
    set_value(CameraSettings.ShutterTimingExposureTime, ms)
    return get_exp_time_ms()

def get_exp_time_ms():
    return experiment.GetValue(CameraSettings.ShutterTimingExposureTime)

def set_cwl(cwl):
    set_value(SpectrometerSettings.GratingCenterWavelength, cwl)
    return get_cwl()

def get_cwl():
    return experiment.GetValue(SpectrometerSettings.GratingCenterWavelength)

def set_grating(g):
    if g in get_available_gratings():
        experiment.SetValue(SpectrometerSettings.GratingSelected,g) #g = '[500nm,1800][0][0]', '[500nm,600][1][0]', '[500nm,300][2][0]'
        return get_grating()
    else:
        print('Invalid input %s. Use one of the values: %s'%(g, str(get_available_gratings())))

def get_grating():
    return experiment.GetValue(SpectrometerSettings.GratingSelected)

def get_available_gratings():
    gs = list(experiment.GetCurrentCapabilities(SpectrometerSettings.GratingSelected))
    return gs

def update_height_width():
    roi = experiment.SelectedRegions[0]
    global _h, _w
    _h = roi.get_Height()
    _w = roi.get_Width()
    _ybin = roi.get_YBinning()
    _h /= _ybin
    
def _convert_array1D_to_data(array1D):
    if _h == 1:
        ans = array1D
    else:
        ans = np.reshape(array1D,(_h,_w))
    return ans

def update_save_directory(p):
    experiment.SetValue(ExperimentSettings.FileNameGenerationDirectory, os.path.join(p))
    
def reset_increment_number():
    experiment.SetValue(ExperimentSettings.FileNameGenerationIncrementNumber,0)

def init():
    global auto, experiment, acquireCompleted, _monitoring_list, filemanager
    auto = Automation(True, List[String]())
    experiment = auto.LightFieldApplication.Experiment
    acquireCompleted = AutoResetEvent(False)
    
    experiment.Load('NL_setup2')
    _monitoring_list = List[String]()
    _monitoring_list.Add(String(ExperimentSettings.ResultantFrameSize))
    experiment.FilterSettingChanged(_monitoring_list)
    update_height_width()
    reset_increment_number()
    experiment.ExperimentCompleted += experiment_completed
    experiment.SettingChanged += experiment_setting_changed
    
    filemanager = auto.LightFieldApplication.FileManager

def try_activate_window():
    try:
        auto.ActivateWindow()
    except:
        init()

