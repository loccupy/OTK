
from gurux_dlms import GXDLMSClient
from gurux_dlms.secure import GXDLMSSecureClient
from gurux_common.io import Parity, StopBits, BaudRate
from gurux_serial.GXSerial import GXSerial
from gurux_common.enums import TraceLevel

from gurux_dlms.enums import InterfaceType, Authentication, Security, Standard
from gurux_dlms import GXDLMSClient
from gurux_dlms.secure import GXDLMSSecureClient
from gurux_dlms.GXByteBuffer import GXByteBuffer
from gurux_dlms.objects import GXDLMSObject
from gurux_common.enums import TraceLevel
from gurux_common.io import Parity, StopBits, BaudRate
from gurux_net.enums import NetworkType
from gurux_net import GXNet
from gurux_serial.GXSerial import GXSerial


class GXSettings:
    def __init__(self):
        self.media = None
        self.trace = TraceLevel.INFO
        self.invocationCounter = None
        self.client = GXDLMSSecureClient(True)

        self.readObjects = []
        # self.outputFile = None
        self.outputFile = 'E:\\Python\\prom_energo\\gurux_dlms\\libs\\123'

    def getParameters(self, interface_type: str, port: str, password: str, authentication: str, serverAddress: int,
                      logicalAddress: int, clientAddress: int, baudRate: int):
        if interface_type == "COM":
            self.media = GXSerial(None)
            self.media.port = port

        self.media.baudRate = baudRate  # -S
        self.media.dataBits = 8
        self.media.parity = Parity.NONE
        self.media.stopbits = StopBits.ONE

        if authentication == "None":
            self.client.authentication = Authentication.NONE  # -a   (None, Low, High)
        elif authentication == "Low":
            self.client.authentication = Authentication.LOW
        elif authentication == "High":
            self.client.authentication = Authentication.HIGH

        self.client.password = password  # -P
        self.client.clientAddress = clientAddress  # -c
        self.client.serverAddress = serverAddress  # -s
        self.client.serverAddress = GXDLMSClient.getServerAddress(logicalAddress, self.client.serverAddress)  # -l
        self.client.limits.maxInfoRX = 114
        self.client.limits.maxInfoTX = 114
