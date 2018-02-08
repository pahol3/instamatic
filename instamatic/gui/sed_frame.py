from tkinter import *
from tkinter.ttk import *
import tkinter.messagebox
import os, sys

import matplotlib
matplotlib.use('TkAgg')
from instamatic.calibrate import CalibBeamShift, CalibDirectBeam
from instamatic.calibrate.filenames import CALIB_DIRECTBEAM, CALIB_BEAMSHIFT

message1 = """
 1. Go to diffraction mode and select desired camera length (CAM L)
 2. Center the beam with diffraction shift (PLA)
 3. Focus the diffraction pattern (DIFF FOCUS)
    
Press <OK> to start"""

message2 = """
 1. Go to desired magnification (e.g. 2500x)
 2. Select desired beam size (BRIGHTNESS)
 3. Center the beam with beamshift
    
Press <OK> to start"""

message3 = """
 1. Move the stage to where you want to begin
 2. Follow the instructions in the terminal

Press <OK> to start"""


class ExperimentalSED(LabelFrame):
    """docstring for ExperimentalSED"""
    def __init__(self, parent):
        LabelFrame.__init__(self, parent, text="Serial electron diffraction")
        self.parent = parent

        self.calib_path = ""

        self.init_vars()

        frame = Frame(self)

        Label(frame, text="Scan area (um)").grid(row=5, column=0, sticky="W")
        self.e_start_x = Entry(frame, width=20, textvariable=self.var_scan_radius)
        self.e_start_x.grid(row=5, column=1, padx=10)

        Label(frame, text="Exp. time image:").grid(row=6, column=0, sticky="W")
        self.e_exp_time_image = Entry(frame, width=20, textvariable=self.var_image_exposure)
        self.e_exp_time_image.grid(row=6, column=1, padx=10)

        Label(frame, text="Exp. time diff:").grid(row=7, column=0, sticky="W")
        self.e_exp_time_diff = Entry(frame, width=20, textvariable=self.var_diff_exposure)
        self.e_exp_time_diff.grid(row=7, column=1, padx=10)

        Label(frame, text="Brightness:").grid(row=6, column=2, sticky="W")
        self.e_exp_time_diff = Entry(frame, width=20, textvariable=self.var_diff_brightness)
        self.e_exp_time_diff.grid(row=6, column=3, padx=10)

        Label(frame, text="Spot size:").grid(row=7, column=2, sticky="W")
        self.e_exp_time_diff = Entry(frame, width=20, textvariable=self.var_image_spotsize)
        self.e_exp_time_diff.grid(row=7, column=3, padx=10)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(2, weight=1)

        frame.pack(side="top", fill="x", padx=10, pady=10)

        frame = Frame(self)        
        self.ShowCalibBeamshift = Button(frame, text="Show calib beamshift", command=self.show_calib_beamshift, state=NORMAL)
        self.ShowCalibBeamshift.grid(row=1, column=0, sticky="EW")

        self.ShowCalibDirectBeam1 = Button(frame, text="Show calib directbeam1", command=self.show_calib_directbeam1, state=NORMAL)
        self.ShowCalibDirectBeam1.grid(row=1, column=1, sticky="EW")

        self.ShowCalibDirectBeam2 = Button(frame, text="Show calib directbeam2", command=self.show_calib_directbeam2, state=NORMAL)
        self.ShowCalibDirectBeam2.grid(row=1, column=2, sticky="EW")
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        frame.columnconfigure(2, weight=1)
        frame.pack(side="top", fill="both", expand=True, padx=10, pady=10)

        frame = Frame(self)

        self.CollectionButton = Button(frame, text="Start Collection", command=self.start_collection, state=NORMAL)
        self.CollectionButton.pack(side="bottom", fill="both")

        frame.pack(side="bottom", fill="both", padx=10, pady=10)

    def init_vars(self):
        self.var_scan_radius = DoubleVar(value=100)
        self.var_image_exposure = DoubleVar(value=0.5)
        self.var_diff_exposure = DoubleVar(value=0.1)
        self.var_image_spotsize = IntVar(value=4)
        self.var_diff_brightness = IntVar(value=40000)

    def set_trigger(self, trigger=None, q=None):
        self.triggerEvent = trigger
        self.q = q

    def start_collection(self):
        okay = tkinter.messagebox.askokcancel("Start experiment", message3, icon='warning')
        if okay:
            params = self.get_params()
            self.q.put(("sed", params))
            self.triggerEvent.set()

    def show_calib_beamshift(self):
        # TODO: use mpl_frame.ShowMatplotlibFig
        path = os.path.join(self.calib_path, CALIB_BEAMSHIFT)
        c = CalibDirectBeam.from_file(path)
        c.plot()

    def show_calib_directbeam1(self):
        # TODO: use mpl_frame.ShowMatplotlibFig
        path = os.path.join(self.calib_path, CALIB_DIRECTBEAM)
        c = CalibDirectBeam.from_file(path)
        c.plot("DiffShift")

    def show_calib_directbeam2(self):
        # TODO: use mpl_frame.ShowMatplotlibFig
        path = os.path.join(self.calib_path, CALIB_DIRECTBEAM)
        c = CalibDirectBeam.from_file(path)
        c.plot("BeamShift")

    def get_params(self):
        params = { "image_exposure": self.var_image_exposure.get(),
                   "image_spotsize": self.var_image_spotsize.get(),
                   "diff_exposure": self.var_diff_exposure.get(),
                   "diff_spotsize": self.var_image_spotsize.get(),
                   "diff_brightness": self.var_diff_brightness.get(),
                   "scan_radius": self.var_scan_radius.get() }
        return params


if __name__ == '__main__':
    root = Tk()
    ExperimentalSED(root).pack(side="top", fill="both", expand=True)
    root.mainloop()

