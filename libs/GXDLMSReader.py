#
#  --------------------------------------------------------------------------
#   Gurux Ltd
#
#
#
#  Filename: $HeadURL$
#
#  Version: $Revision$,
#                   $Date$
#                   $Author$
#
#  Copyright (c) Gurux Ltd
#
# ---------------------------------------------------------------------------
#
#   DESCRIPTION
#
#  This file is a part of Gurux Device Framework.
#
#  Gurux Device Framework is Open Source software; you can redistribute it
#  and/or modify it under the terms of the GNU General Public License
#  as published by the Free Software Foundation; version 2 of the License.
#  Gurux Device Framework is distributed in the hope that it will be useful,
#  but WITHOUT ANY WARRANTY; without even the implied warranty of
#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
#  See the GNU General Public License for more details.
#
#  More information of Gurux products: http://www.gurux.org
#
#  This code is licensed under the GNU General Public License v2.
#  Full text may be retrieved at http://www.gnu.org/licenses/gpl-2.0.txt
# ---------------------------------------------------------------------------
import os
import datetime
import time
import traceback
import random
from gurux_common.enums import TraceLevel
from gurux_common.io import Parity, StopBits
from gurux_common import ReceiveParameters, GXCommon, TimeoutException
from gurux_dlms import GXByteBuffer, GXReplyData, GXDLMSTranslator, GXDLMSException, GXDLMSAccessItem, GXDLMSClient,\
    GXTime, GXDateTime
from gurux_dlms.enums import InterfaceType, ObjectType, Authentication, Conformance, DataType,\
    Security, AssociationResult, SourceDiagnostic, AccessServiceCommandType
from gurux_dlms.objects import GXDLMSObject, GXDLMSObjectCollection, GXDLMSData, GXDLMSRegister, \
    GXDLMSDemandRegister, GXDLMSProfileGeneric, GXDLMSExtendedRegister, GXDLMSDisconnectControl, \
    GXDLMSActivityCalendar, GXDLMSDayProfile, GXDLMSWeekProfile, GXDLMSSeasonProfile, GXDLMSDayProfileAction, \
    GXDLMSScriptTable, GXDLMSClock, GXDLMSAssociationLogicalName
from gurux_net import GXNet
from gurux_dlms.objects import GXDLMSImageTransfer#это я написал


class GXDLMSReader(GXDLMSDisconnectControl):
    #pylint: disable=too-many-public-methods, too-many-instance-attributes
    def __init__(self, client, media, trace, invocationCounter):
        super().__init__()
        #pylint: disable=too-many-arguments
        self.replyBuff = bytearray(8 + 1024)
        self.waitTime = 5000
        self.logFile = open("../logFile.txt", "w")
        self.trace = trace
        self.media = media
        self.invocationCounter = invocationCounter
        self.client = client
        if self.trace > TraceLevel.WARNING:
            print("Authentication: " + str(self.client.authentication))
            print("ClientAddress: " + hex(self.client.clientAddress))
            print("ServerAddress: " + hex(self.client.serverAddress))

    def disconnect(self):
        #pylint: disable=broad-except
        if self.media and self.media.isOpen():
            print("DisconnectRequest")
            reply = GXReplyData()
            self.readDLMSPacket(self.client.disconnectRequest(), reply)

    def release(self):
        #pylint: disable=broad-except
        if self.media and self.media.ismedia():
            print("DisconnectRequest")
            reply = GXReplyData()
            try:
                #Release is call only for secured connections.
                #All meters are not supporting Release and it's causing
                #problems.
                if self.client.interfaceType == InterfaceType.WRAPPER or self.client.ciphering.security != Security.NONE:
                    self.readDataBlock(self.client.releaseRequest(), reply)
            except Exception:
                pass
                #  All meters don't support release.

    def close(self):
        #pylint: disable=broad-except
        if self.media and self.media.isOpen():
            print("DisconnectRequest")
            reply = GXReplyData()
            try:
                #Release is call only for secured connections.
                #All meters are not supporting Release and it's causing
                #problems.
                if self.client.interfaceType == InterfaceType.WRAPPER or self.client.ciphering.security != Security.NONE:
                    self.readDataBlock(self.client.releaseRequest(), reply)
            except Exception:
                pass
                #  All meters don't support release.
            reply.clear()
            self.readDLMSPacket(self.client.disconnectRequest(), reply)
            self.media.close()

    @classmethod
    def now(cls):
        return datetime.datetime.now().strftime("%H:%M:%S")

    def writeTrace(self, line, level):
        #print(line)
        self.logFile.write(line + "\n")

    def readDLMSPacket(self, data, reply=None):
        if not reply:
            reply = GXReplyData()
        if isinstance(data, bytearray):
            self.readDLMSPacket2(data, reply)
        elif data:
            for it in data:
                reply.clear()
                self.readDLMSPacket2(it, reply)

    def readDLMSPacket2(self, data, reply):
        if not data:
            return
        notify = GXReplyData()
        reply.error = 0
        eop = 0x7E
        #In network connection terminator is not used.
        if self.client.interfaceType == InterfaceType.WRAPPER and isinstance(self.media, GXNet):
            eop = None
        p = ReceiveParameters()
        p.eop = eop
        p.allData = True
        p.waitTime = self.waitTime
        if eop is None:
            p.Count = 8
        else:
            p.Count = 5
        self.media.eop = eop
        rd = GXByteBuffer()
        with self.media.getSynchronous():
            if not reply.isStreaming():
                self.writeTrace("TX: " + self.now() + "\t" + GXByteBuffer.hex(data), TraceLevel.VERBOSE)
                self.media.send(data)
            pos = 0
            try:
                while not self.client.getData(rd, reply, notify):
                    if notify.data.size != 0:
                        if not notify.isMoreData():
                            t = GXDLMSTranslator()
                            xml = t.dataToXml(notify.data)
                            print(xml)
                            notify.clear()
                        continue
                    if not p.eop:
                        p.count = self.client.getFrameSize(rd)
                    while not self.media.receive(p):
                        pos += 1
                        if pos == 3:
                            raise TimeoutException("Failed to receive reply from the device in given time.")
                        print("Data send failed.  Try to resend " + str(pos) + "/3")
                        self.media.send(data, None)
                    rd.set(p.reply)
                    p.reply = None
            except Exception as e:
                self.writeTrace("RX: " + self.now() + "\t" + str(rd), TraceLevel.ERROR)
                raise e
            self.writeTrace("RX: " + self.now() + "\t" + str(rd), TraceLevel.VERBOSE)
            # print("RX:" + str(rd))
            if reply.error != 0:
                raise GXDLMSException(reply.error)

    def readDataBlock(self, data, reply):
        if data:
            if isinstance(data, (list)):
                for it in data:
                    reply.clear()
                    self.readDataBlock(it, reply)
                return reply.error == 0
            else:
                self.readDLMSPacket(data, reply)
                while reply.isMoreData():
                    if reply.isStreaming():
                        data = None
                    else:
                        data = self.client.receiverReady(reply)
                    self.readDLMSPacket(data, reply)

    def initializeOpticalHead(self):
        if self.client.interfaceType == InterfaceType.HDLC_WITH_MODE_E:
            p = ReceiveParameters()
            p.allData = True
            p.eop = '\n'
            p.waitTime = self.waitTime
            with self.media.getSynchronous():
                data = "/?!\r\n"
                self.writeTrace("TX: " + self.now() + "\t" + data, TraceLevel.VERBOSE)
                self.media.send(data)
                if not self.media.receive(p):
                    raise Exception("Failed to received reply from the media.")

                self.writeTrace("RX: " + self.now() + "\t" + str(p.reply), TraceLevel.VERBOSE)
                #If echo is used.
                if data.encode() == p.reply:
                    p.reply = None
                    if not self.media.receive(p):
                        raise Exception("Failed to received reply from the media.")
                    self.writeTrace("RX: " + self.now() + "\t" + str(p.reply), TraceLevel.VERBOSE)

            if not p.reply or p.reply[0] != ord('/'):
                raise Exception("Invalid responce : " + str(p.reply))
            baudrate = chr(p.reply[4])
            if baudrate == '0':
                bitrate = 300
            elif baudrate == '1':
                bitrate = 600
            elif baudrate == '2':
                bitrate = 1200
            elif baudrate == '3':
                bitrate = 2400
            elif baudrate == '4':
                bitrate = 4800
            elif baudrate == '5':
                bitrate = 9600
            elif baudrate == '6':
                bitrate = 19200
            else:
                raise Exception("Unknown baud rate.")

            print("Bitrate is : " + str(bitrate))
            #Send ACK
            #Send Protocol control character
            controlCharacter = ord('2')
            #"2" HDLC protocol procedure (Mode E)
            #Mode control character
            #"2" //(HDLC protocol procedure) (Binary mode)
            modeControlCharacter = ord('2')
            #Set mode E.
            tmp = bytearray([0x06, controlCharacter, ord(baudrate), modeControlCharacter, 13, 10])
            p.reply = None
            with self.media.getSynchronous():
                self.media.send(tmp)
                #This sleep make sure that all meters can be read.
                time.sleep(1)
                self.writeTrace("TX: " + self.now() + "\t" + GXCommon.toHex(tmp), TraceLevel.VERBOSE)
                p.waitTime = 200
                if self.media.receive(p):
                    self.writeTrace("RX: " + self.now() + "\t" + str(p.reply), TraceLevel.VERBOSE)
                self.media.dataBits = 8
                self.media.parity = Parity.NONE
                self.media.stopBits = StopBits.ONE
                self.media.baudRate = bitrate
                #This sleep make sure that all meters can be read.
                time.sleep(1)

    def updateFrameCounter(self):
        if self.invocationCounter and self.client.ciphering is not None and self.client.ciphering.security != Security.NONE:
            self.initializeOpticalHead()
            self.client.proposedConformance |= Conformance.GENERAL_PROTECTION
            add = self.client.clientAddress
            auth = self.client.authentication
            security = self.client.ciphering.security
            challenge = self.client.ctoSChallenge
            try:
                self.client.clientAddress = 16
                self.client.authentication = Authentication.NONE
                self.client.ciphering.security = Security.NONE
                reply = GXReplyData()
                data = self.client.snrmRequest()
                if data:
                    self.readDLMSPacket(data, reply)
                    self.client.parseUAResponse(reply.data)
                    size = self.client.hdlcSettings.maxInfoTX # + 40
                    self.replyBuff = bytearray(size)
                reply.clear()
                self.readDataBlock(self.client.aarqRequest(), reply)
                self.client.parseAareResponse(reply.data)
                reply.clear()
                d = GXDLMSData(self.invocationCounter)
                self.read(d, 2)
                self.client.ciphering.invocationCounter = 1 + d.value
                print("Invocation counter: " + str(self.client.ciphering.invocationCounter))
                self.disconnect()
                #except Exception as ex:
            finally:
                self.client.clientAddress = add
                self.client.authentication = auth
                self.client.ciphering.security = security
                self.client.ctoSChallenge = challenge

    def initializeConnection(self):
        print("Standard: " + str(self.client.standard))
        if self.client.ciphering.security != Security.NONE:
            print("Security: " + str(self.client.ciphering.security))
            print("System title: " + GXCommon.toHex(self.client.ciphering.systemTitle))
            print("Authentication key: " + GXCommon.toHex(self.client.ciphering.authenticationKey))
            print("Block cipher key: " + GXCommon.toHex(self.client.ciphering.blockCipherKey))
            if self.client.ciphering.dedicatedKey:
                print("Dedicated key: " + GXCommon.toHex(self.client.ciphering.dedicatedKey))
        self.updateFrameCounter()
        self.initializeOpticalHead()
        reply = GXReplyData()
        data = self.client.snrmRequest()
        if data:
            self.readDLMSPacket(data, reply)
            self.client.parseUAResponse(reply.data)
            size = self.client.hdlcSettings.maxInfoTX #+ 40
            self.replyBuff = bytearray(size)
        reply.clear()
        self.readDataBlock(self.client.aarqRequest(), reply)
        self.client.parseAareResponse(reply.data)
        reply.clear()
        if self.client.authentication > Authentication.LOW:
            try:
                for it in self.client.getApplicationAssociationRequest():
                    self.readDLMSPacket(it, reply)
                self.client.parseApplicationAssociationResponse(reply.data)
            except GXDLMSException as ex:
                #Invalid password.
                raise GXDLMSException(AssociationResult.PERMANENT_REJECTED, SourceDiagnostic.AUTHENTICATION_FAILURE)

    def read(self, item, attributeIndex):
        data = self.client.read(item, attributeIndex)[0]
        reply = GXReplyData()
        self.readDataBlock(data, reply)
        #Update data type on read.
        if item.getDataType(attributeIndex) == DataType.NONE:
            item.setDataType(attributeIndex, reply.valueType)
        return self.client.updateValue(item, attributeIndex, reply.value)

    def read_week_profile(self, item, attributeIndex):
        data = self.client.read(item, attributeIndex)[0]
        reply = GXReplyData()
        self.readDataBlock(data, reply)
        data = reply.value
        for _ in data:
            _[0] = _[0].decode('utf-8')
            _[1] = _[1:]
            del _[2:]
        return data

    def _season_parser(self, gx_time):
        if len(gx_time) == 24:
            gx_time = gx_time[:-7]

        time_list = list(gx_time)
        time_list[0], time_list[1], time_list[3], time_list[4] = time_list[3], time_list[4], time_list[0], time_list[1]

        if len(time_list) == 17:
            time_list.insert(6, '20')
        elif len(time_list) == 14:
            time_list.insert(5, str(f'/{datetime.datetime.today().year}'))
            # print(time_list)

        return ''.join(time_list)

    def read_season_profile(self, item, attributeIndex):
        data = self.client.read(item, attributeIndex)[0]
        reply = GXReplyData()
        self.readDataBlock(data, reply)
        data = reply.value
        time_ = self.client.updateValue(item, attributeIndex, reply.value)
        time_ = [str(time_[i]).partition(' ') for i in range(len(time_))]
        for i, _ in enumerate(data):
            _[0] = _[0].decode('utf-8')
            _[1] = self._season_parser(time_[i][-1])
            _[2] = _[2].decode('utf-8')
        return data


    def readList(self, list_):
        if list_:
            data = self.client.readList(list_)
            reply = GXReplyData()
            values = list()
            for it in data:
                self.readDataBlock(it, reply)
                if reply.value:
                    values.extend(reply.value)
                reply.clear()
            if len(values) != len(list_):
                raise ValueError("Invalid reply. Read items count do not match.")
            self.client.updateValues(list_, values)

    def write(self, item, attributeIndex):
        data = self.client.write(item, attributeIndex)
        self.readDLMSPacket(data)

    def readRowsByEntry(self, pg, index, count):
        data = self.client.readRowsByEntry(pg, index, count)
        reply = GXReplyData()
        self.readDataBlock(data, reply)
        return self.client.updateValue(pg, 2, reply.value)

    def readRowsByRange(self, pg, start, end):
        reply = GXReplyData()
        data = self.client.readRowsByRange(pg, start, end)
        self.readDataBlock(data, reply)
        return self.client.updateValue(pg, 2, reply.value)

    #Read values using Access request.
    def readByAccess(self, list_):
        if list_:
            reply = GXReplyData()
            data = self.client.accessRequest(None, list_)
            self.readDataBlock(data, reply)
            self.client.parseAccessResponse(list_, reply.data)

    def readScalerAndUnits(self):
        #pylint: disable=broad-except
        objs = self.client.objects.getObjects([ObjectType.REGISTER, ObjectType.EXTENDED_REGISTER, ObjectType.DEMAND_REGISTER])
        list_ = list()
        if self.client.negotiatedConformance & Conformance.ACCESS != 0:
            for it in objs:
                if isinstance(it, (GXDLMSRegister, GXDLMSExtendedRegister)):
                    list_.append(GXDLMSAccessItem(AccessServiceCommandType.GET, it, 3))
                elif isinstance(it, (GXDLMSDemandRegister)):
                    list_.append(GXDLMSAccessItem(AccessServiceCommandType.GET, it, 4))
            self.readByAccess(list_)
            return
        try:
            if self.client.negotiatedConformance & Conformance.MULTIPLE_REFERENCES != 0:
                for it in objs:
                    if isinstance(it, (GXDLMSRegister, GXDLMSExtendedRegister)):
                        list_.append((it, 3))
                    elif isinstance(it, (GXDLMSDemandRegister,)):
                        list_.append((it, 4))
                self.readList(list_)
        except Exception:
            self.client.negotiatedConformance &= ~Conformance.MULTIPLE_REFERENCES
        if self.client.negotiatedConformance & Conformance.MULTIPLE_REFERENCES == 0:
            for it in objs:
                try:
                    if isinstance(it, (GXDLMSRegister,)):
                        self.read(it, 3)
                    elif isinstance(it, (GXDLMSDemandRegister,)):
                        self.read(it, 4)
                except Exception:
                    pass

    def getProfileGenericColumns(self):
        #pylint: disable=broad-except
        profileGenerics = self.client.objects.getObjects(ObjectType.PROFILE_GENERIC)
        for pg in profileGenerics:
            self.writeTrace("Profile Generic " + str(pg.name) + "Columns:", TraceLevel.INFO)
            try:
                if pg.canRead(3):
                    self.read(pg, 3)
                if self.trace > TraceLevel.WARNING:
                    sb = ""
                    for k, _ in pg.captureObjects:
                        if sb:
                            sb += " | "
                        sb += str(k.name)
                        sb += " "
                        desc = k.description
                        if desc:
                            sb += desc
                    self.writeTrace(sb, TraceLevel.INFO)
            except Exception as ex:
                self.writeTrace("Err! Failed to read columns:" + str(ex), TraceLevel.ERROR)

    def getReadOut(self):
        #pylint: disable=unidiomatic-typecheck, broad-except
        for it in self.client.objects:
            if type(it) == GXDLMSObject:
                print("Unknown Interface: " + it.objectType.__str__())
                continue
            if isinstance(it, GXDLMSProfileGeneric):
                continue

            self.writeTrace("-------- Reading " + str(it.objectType) + " " + str(it.name) + " " + it.description, TraceLevel.INFO)
            for pos in it.getAttributeIndexToRead(True):
                try:
                    if it.canRead(pos):
                        val = self.read(it, pos)
                        self.showValue(pos, val)
                    else:
                        self.writeTrace("Attribute" + str(pos) + " is not readable.", TraceLevel.INFO)
                except Exception as ex:
                    self.writeTrace("Error! Index: " + str(pos) + " " + str(ex), TraceLevel.ERROR)
                    self.writeTrace(str(ex), TraceLevel.ERROR)
                    if not isinstance(ex, (GXDLMSException, TimeoutException)):
                        traceback.print_exc()

    def showValue(self, pos, val):
        if isinstance(val, (bytes, bytearray)):
            val = GXByteBuffer(val)
        elif isinstance(val, list):
            str_ = ""
            for tmp in val:
                if str_:
                    str_ += ", "
                if isinstance(tmp, bytes):
                    str_ += GXByteBuffer.hex(tmp)
                else:
                    str_ += str(tmp)
            val = str_
        self.writeTrace("Index: " + str(pos) + " Value: " + str(val), TraceLevel.INFO)

    def getProfileGenerics(self):
        #pylint: disable=broad-except,too-many-nested-blocks
        cells = []
        profileGenerics = self.client.objects.getObjects(ObjectType.PROFILE_GENERIC)
        for it in profileGenerics:
            self.writeTrace("-------- Reading " + str(it.objectType) + " " + str(it.name) + " " + it.description, TraceLevel.INFO)
            entriesInUse = self.read(it, 7)
            entries = self.read(it, 8)
            self.writeTrace("Entries: " + str(entriesInUse) + "/" + str(entries), TraceLevel.INFO)
            pg = it
            if entriesInUse == 0 or not pg.captureObjects:
                continue
            try:
                cells = self.readRowsByEntry(pg, 1, 1)
                if self.trace > TraceLevel.WARNING:
                    for rows in cells:
                        for cell in rows:
                            if isinstance(cell, bytearray):
                                self.writeTrace(GXByteBuffer.hex(cell) + " | ", TraceLevel.INFO)
                            else:
                                self.writeTrace(str(cell) + " | ", TraceLevel.INFO)
                        self.writeTrace("", TraceLevel.INFO)
            except Exception as ex:
                self.writeTrace("Error! Failed to read first row: " + str(ex), TraceLevel.ERROR)
                if not isinstance(ex, (GXDLMSException, TimeoutException)):
                    traceback.print_exc()
            try:
                start = datetime.datetime.now()
                end = start
                start = start.replace(hour=0, minute=0, second=0, microsecond=0)
                end = end.replace(minute=0, second=0, microsecond=0)
                cells = self.readRowsByRange(it, start, end)
                for rows in cells:
                    row = ""
                    for cell in rows:
                        if row:
                            row += " | "
                        if isinstance(cell, bytearray):
                            row += GXByteBuffer.hex(cell)
                        else:
                            row += str(cell) 
                    self.writeTrace(row, TraceLevel.INFO)
            except Exception as ex:
                self.writeTrace("Error! Failed to read last day: " + str(ex), TraceLevel.ERROR)

    def getAssociationView(self):
        reply = GXReplyData()
        self.readDataBlock(self.client.getObjectsRequest(), reply)
        self.client.parseObjects(reply.data, True, False)
        #Access rights must read differently when short Name referencing is used.
        if not self.client.useLogicalNameReferencing:
            sn = self.client.objects.findBySN(0xFA00)
            if sn and sn.version > 0:
                try:
                    self.read(sn, 3)
                except (GXDLMSException):
                    self.writeTrace("Access rights are not implemented for the meter.", TraceLevel.INFO)

    def readAll(self, outputFile):
        try:
            read = False
            self.initializeConnection()


            if outputFile and os.path.exists(outputFile):
                try:
                    c = GXDLMSObjectCollection.load(outputFile)
                    self.client.objects.extend(c)
                    if self.client.objects:
                        read = True
                except Exception:
                    read = False
            if not read:

                print("Вот эта штука собирает коллекцию")
                self.getAssociationView()
                self.readScalerAndUnits()
                self.getProfileGenericColumns()
            self.getReadOut()
            self.getProfileGenerics()
            if outputFile:
                self.client.objects.save(outputFile)
        except (KeyboardInterrupt, SystemExit):
            #Don't send anything if user is closing the app.
            self.media = None
            raise
        finally:
            self.close()

    def relay_actions(self, actions):
        reply = GXReplyData()
        dc = GXDLMSDisconnectControl("0.0.96.3.10.255")
        if actions == 0:
            self.readDataBlock(dc.remoteDisconnect(self.client), reply)
        elif actions == 1:
            self.readDataBlock(dc.remoteReconnect(self.client), reply)

    def relay_disconnect(self):
        self.relay_actions(0)

    def relay_reconnect(self):
        self.relay_actions(1)

    # convert date and time accounting difference 2 hours 59 munites 1 second
    def normalize_time(self, write_time):  # преобразует время формата datetime, вычитая 2 часа 59 минут 1 секунду
        write_time = write_time - datetime.timedelta(hours=2, minutes=59, seconds=1)
        return write_time

    def check_period_profile(self, pg):
        self.read(pg, 3)
        capture_period = self.read(pg, 4)
        buffer = self.read(pg, 2)
        if not buffer:
            return 0, 0
        delta = datetime.timedelta(hours=2, minutes=59)
        for _ in buffer:
            _[0] = datetime.datetime.strptime(str(_[0]), '%m/%d/%y %H:%M:%S') - delta
        first_record = buffer[0][0]
        last_record = buffer[-1][0]
        time_diff = int((last_record - first_record).total_seconds() // capture_period) + 1
        return len(buffer), time_diff


    # convert date and time to tuple for further write
    def convert_date_time_to_tuple(self, date_time):
        day = int(date_time[:2])
        month = int(date_time[3:5])
        year = int(date_time[6:10])
        hours = int(date_time[11:13])
        minutes = int(date_time[14:16])
        seconds = int(date_time[17:])
        return year, month, day, hours, minutes, seconds

    # convert time (e.g. tariff interval start) to tuple(hours, min, sec)
    def convert_time_to_tuple(self, time):
        hours = int(time[:2])
        minutes = int(time[3:5])
        seconds = int(time[6:])
        return hours, minutes, seconds

    # convert list(hours, min, sec) to standard time view - HH:MM:SS
    def convert_list_to_time(self, lst):
        wr_time = ''
        if lst[0] < 10:
            wr_time += '0' + str(lst[0]) + ':'
        else:
            wr_time += str(lst[0]) + ':'
        if lst[1] < 10:
            wr_time += '0' + str(lst[1]) + ':'
        else:
            wr_time += str(lst[1]) + ':'
        wr_time += '00'
        return wr_time

    def convert_list_to_datetime(self, lst):
        wr_datetime = ''
        if lst[2] < 10:
            wr_datetime += '0' + str(lst[2]) + '/'
        else:
            wr_datetime += str(lst[2]) + '/'
        if lst[1] < 10:
            wr_datetime += '0' + str(lst[1]) + '/'
        else:
            wr_datetime += str(lst[1]) + '/'
        wr_datetime += str(lst[0])
        wr_datetime += ' 00:00:00'
        return wr_datetime

    # check whether datetime format is correct
    def check_datetime_format(self, date_time):
        status = False
        day = int(date_time[3:5])
        month = int(date_time[:2])
        year = int(date_time[6:8])
        hour = int(date_time[9:11])
        minute = int(date_time[12:14])
        second = int(date_time[15:])
        if len(date_time) == 17 and (0 < day < 32) and (0 < month < 13) and (year < 2030) and (hour < 24) and (
                minute < 60) and (second < 60):
            status = True
        return status

    # add tariff interval in daySchedule of dayID
    def add_tariff_interval(self, day_schedule, start_time, selector):
        day_schedule.startTime = GXTime(datetime.datetime(2022, 1, 1, start_time[0], start_time[1], start_time[2]))
        day_schedule.scriptLogicalName = '0.0.10.0.100.255'
        day_schedule.scriptSelector = selector

    # add day schedule (some number of tariff intervals) for dayID
    #     and return that day schedule as a list
    def add_day_schedule(self, written_calendar, day, interval_count):
        day.daySchedules = list()
        written_calendar.append(list())
        tariff_intervals = [[0, 0, 0], [1, 30, 0], [2, 0, 0], [6, 0, 0], [6, 30, 0], [10, 15, 0], [11, 0, 0],
                            [22, 0, 0]]
        selectors = [1, 2, 3, 4]
        for i in range(interval_count):
            day.daySchedules.append(GXDLMSDayProfileAction())
            selector = random.choice(selectors)
            start_time = self.convert_list_to_time(tariff_intervals[i])
            written_calendar[1].append(start_time)
            written_calendar[1].append('0.0.10.0.100.255')
            written_calendar[1].append(selector)
            self.add_tariff_interval(day.daySchedules[i], tariff_intervals[i], selector)
        return written_calendar

    # add day profile (some numbers of day ID) and return that day profile as a list
    def add_day_profile(self, written_calendar, calendar, day_count, interval_count):
        calendar.dayProfileTablePassive = list()
        day_ids = [1, 2, 3, 4]
        for i in range(day_count):
            written_calendar.append(list())
            written_calendar[i].append(day_ids[i])
            calendar.dayProfileTablePassive.append(GXDLMSDayProfile())
            calendar.dayProfileTablePassive[i].dayId = day_ids[i]
            self.add_day_schedule(written_calendar[i], calendar.dayProfileTablePassive[i], interval_count)
        return written_calendar

    def to_ascii(self, name):
        s = ''
        for i in name:
            s += str(hex(ord(i)))[2:]
        return s

    # add new week to week profile
    def add_week(self, written_calendar, week_schedule, week_name):
        written_calendar.append(list())
        week_schedule.name = self.to_ascii(week_name)
        week_schedule.monday = random.choice(range(1, 5))
        written_calendar[1].append(week_schedule.monday)
        week_schedule.tuesday = random.choice(range(1, 5))
        written_calendar[1].append(week_schedule.tuesday)
        week_schedule.wednesday = random.choice(range(1, 5))
        written_calendar[1].append(week_schedule.wednesday)
        week_schedule.thursday = random.choice(range(1, 5))
        written_calendar[1].append(week_schedule.thursday)
        week_schedule.friday = random.choice(range(1, 5))
        written_calendar[1].append(week_schedule.friday)
        week_schedule.saturday = random.choice(range(1, 5))
        written_calendar[1].append(week_schedule.saturday)
        week_schedule.sunday = random.choice(range(1, 5))
        written_calendar[1].append(week_schedule.sunday)
        return written_calendar

    # add full week profile
    def add_week_profile(self, written_calendar, calendar, week_count):
        calendar.weekProfileTablePassive = list()
        week_names = ['Default', 'week23456week', 'week34567week12', 'w']
        for i in range(week_count):
            written_calendar.append(list())
            written_calendar[i].append(week_names[i])
            calendar.weekProfileTablePassive.append(GXDLMSWeekProfile())
            self.add_week(written_calendar[i], calendar.weekProfileTablePassive[i], week_names[i])
        return written_calendar

    # add week with specified week name and week schedule
    def add_special_week_profile(self, calendar, week_name, week_schedule):
        calendar.weekProfileTablePassive = list()
        calendar.weekProfileTablePassive.append(GXDLMSWeekProfile())
        calendar.weekProfileTablePassive[0].name = week_name
        calendar.weekProfileTablePassive[0].monday = week_schedule[0]
        calendar.weekProfileTablePassive[0].tuesday = week_schedule[1]
        calendar.weekProfileTablePassive[0].wednesday = week_schedule[2]
        calendar.weekProfileTablePassive[0].thursday = week_schedule[3]
        calendar.weekProfileTablePassive[0].friday = week_schedule[4]
        calendar.weekProfileTablePassive[0].saturday = week_schedule[5]
        calendar.weekProfileTablePassive[0].sunday = week_schedule[6]


    # convert weekday name to his number of the week (-1)
    def weekday_to_day_number(self, weekday):
        if weekday == "Monday":
            n = 0
        elif weekday == "Tuesday":
            n = 1
        elif weekday == "Wednesday":
            n = 2
        elif weekday == "Thursday":
            n = 3
        elif weekday == "Friday":
            n = 4
        elif weekday == "Saturday":
            n = 5
        elif weekday == "Sunday":
            n = 6
        return n


    def add_season(self, written_calendar, i, season_schedule, season_name, week_name):
        start_times = [[2022, 1, 1, 0, 0, 0], [2022, 4, 14, 0, 0, 0], [2022, 7, 7, 0, 0, 0], [2022, 10, 31, 7, 0, 0, 0]]
        season_schedule.name = self.to_ascii(season_name)
        season_schedule.weekName = self.to_ascii(week_name)
        season_schedule.start = GXDateTime(datetime.datetime(start_times[i][0], start_times[i][1], start_times[i][2], start_times[i][3],
                                           start_times[i][4], start_times[i][5]))
        written_calendar.append(season_name)
        written_calendar.append(self.convert_list_to_datetime(start_times[i]))
        written_calendar.append(week_name)
        return written_calendar

    def add_season_profile(self, written_calendar, calendar, season_count, weeks):
        calendar.seasonProfilePassive = list()
        season_names = ['Season1', 'Season2', 'Season26464623a', 'Season4']
        for i in range(season_count):
            written_calendar.append(list())
            calendar.seasonProfilePassive.append(GXDLMSSeasonProfile())
            self.add_season(written_calendar[i], i, calendar.seasonProfilePassive[i], season_names[i], weeks[i])
        return written_calendar

    day_id = None
    start_time = None
    script = None
    selector = None

    def _full_time(self, time):
        """ Вспомогательная функция для _read_elem_in_list """
        if len(time) == 1:
            return "0" + time
        else:
            return time

    def _read_elem_in_list(self, data):
        """ Парсит time, script, selector """
        list_ = []
        for elem in data:
            start_time_hour = self._full_time(str(elem[0][0]))
            start_time_minutes = self._full_time(str(elem[0][1]))
            start_time_second = self._full_time(str(elem[0][2]))
            start_time = f"{start_time_hour}:{start_time_minutes}:{start_time_second}"
            script_ = bytes(elem[1])
            script_ = '.'.join([str(el) for el in script_])
            selector_ = elem[2]
            list_.append(start_time)
            list_.append(script_)
            list_.append(selector_)
        return list_

    def _parser_for_day_profile(self, data):
        list_ = []
        """ Если один day id - готово """
        if isinstance(data[0], int):
            day_id = data[0]
            list_.append(day_id)
            data = self._read_elem_in_list(data[1])
            list_.append(data)
            return list_
        else:
            for elem in data:
                data = self._parser_for_day_profile(elem)
                list_.append(data)
            return list_

    def read_activity_calendar(self, item, attributeIndex):
        data = self.client.read(item, attributeIndex)[0]
        reply = GXReplyData()
        self.readDataBlock(data, reply)
        return reply.value

# -------------------------------------------------- Activity Calendar -------------------------------------------------

    def read_day_profile_active(self):
        ac = GXDLMSActivityCalendar()
        return self._parser_for_day_profile(self.read_activity_calendar(ac, 5))

    def read_day_profile_passive(self):
        ac = GXDLMSActivityCalendar()
        return self._parser_for_day_profile(self.read_activity_calendar(ac, 9))

    # def read_day_profile1(self):
    #     from test_settings import settings
    #     ac = GXDLMSActivityCalendar()
    #     ac.index = 9
    #     return ac.getValue(settings, ac)

    # def write_day_profile(self):
    #     from test_settings import settings
    #     data = GXDLMSClient.createObject(ObjectType.ACTIVITY_CALENDAR)
    #     day_profile = read_day_profile_passive()
    #     data.logicalName = "0.0.13.0.0.255"
    #     #data.dayProfileTablePassive =
    #     ac = GXDLMSActivityCalendar()
    #     ac.index = 9
    #     ac.dayProfileTablePassive = "DayProfileTablePassive"
    #     return ac.getValue(settings=settings, e=ac)

    def activate_passive_calendar(self):
        reply = GXReplyData()
        ac = GXDLMSActivityCalendar()
        self.readDataBlock(ac.activatePassiveCalendar(self.client), reply)

    def script_table(self):
        ac = GXDLMSScriptTable(ln="0.0.10.0.100.255")
        data = ac.getValues()
        return data

    # def profile(self):
    #     pg = GXDLMSProfileGeneric(ln="1.0.99.1.0.255")

# --------------------------------------------------- Profile Generic --------------------------------------------------
    def capture_profile(self, obis):
        reply = GXReplyData()
        dc = GXDLMSProfileGeneric(obis)
        self.readDataBlock(dc.capture(self.client), reply)

    def capture_month_profile(self):
        self.capture_profile("1.0.98.1.0.255")

    def capture_day_profile(self):
        self.capture_profile("1.0.98.2.0.255")

    def capture_hour_profile(self):
        self.capture_profile("1.0.99.2.0.255")

    def capture_load_profile(self):
        self.capture_profile("1.0.99.1.0.255")

    def capture_artur(self):
        self.capture_profile("1.0.99.164.0.255")

    def reset_profile(self, obis):
        reply = GXReplyData()
        dc = GXDLMSProfileGeneric(obis)
        self.readDataBlock(dc.reset(self.client), reply)

    def reset_month_profile(self):
        self.reset_profile("1.0.98.1.0.255")

    def reset_day_profile(self):
        self.reset_profile("1.0.98.2.0.255")

    def reset_hour_profile(self):
        self.reset_profile("1.0.99.2.0.255")

    def reset_load_profile(self):
        self.reset_profile("1.0.99.1.0.255")




    #пишу всякую херь

    def my_verify(self, obis):
        reply = GXReplyData()
        dc = GXDLMSImageTransfer(obis)
        self.readDataBlock(dc.imageVerify(self.client), reply)

    def my_activate(self, obis):
        reply = GXReplyData()
        dc = GXDLMSImageTransfer(obis)
        self.readDataBlock(dc.imageActivate(self.client), reply)
        
    def my_shifttime(self, obis,delta=1):
        reply = GXReplyData()
        dc = GXDLMSClock(obis)
        self.readDataBlock(dc.shiftTime(self.client,forTime=delta), reply)

    def my_transfer(self, obis):
        reply = GXReplyData()
        dc = GXDLMSImageTransfer(obis)
        self.readDataBlock(dc.imageTransferInitiate(self.client,imageIdentifier="123",forImageSize=1), reply)

    def my_updatesecret(self, obis):
        reply = GXReplyData()
        dc = GXDLMSAssociationLogicalName(obis)
        self.readDataBlock(dc.updateSecret(self.client,), reply)
