# import canmatrix.formats.ldf
from openpyxl import load_workbook
import pandas as pd
from ldfparser import LDF, LinFrame, LinSignal, LinSignalEncodingType, LinUnconditionalFrame
import canmatrix
import lin
from ctypes import *
from lin.interfaces.peak import PLinApi, LinBus
from lin.interfaces.peak.PLinApi import TLINVersion
from ldfparser import LDF
from ldfparser.frame import LinUnconditionalFrame
from ldfparser.signal import LinSignal
from ldfparser.node import LinMaster, LinSlave
from ldfparser.lin import LinVersion
from ldfparser import save_ldf
from ldfparser.node import LinNode

ldf = LDF()

ldf._protocol_version = LinVersion(2, 1)
ldf._language_version = LinVersion(2, 1)
)

ldf._baudrate = 19200   
ldf._channel = "LIN1"   

master = LinMaster(
    name="Master",
    timebase=0.01,
    jitter=0.001,
    max_header_length=8,
    response_tolerance=0.1
)
slave = LinSlave(name="SlaveNode")

ldf._master = master
ldf._slaves[slave.name] = slave

signal1 = LinSignal(name="AmbientLight", width=1, init_value=0)
signal2 = LinSignal(name="InteriorLight", width=1, init_value=0)

frame = LinUnconditionalFrame(
    frame_id=0x10,
    name="LightControl",
    length=8,
    signals={1: signal1, 2: signal2},
    pad_with_zero=True
)

ldf._unconditional_frames[frame.name] = frame

save_ldf(ldf, "output.ldf", "C:\\projects\\Convert2DBC\\ldf.jinja2")
