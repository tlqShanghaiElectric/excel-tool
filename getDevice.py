# -*- coding: utf-8 -*-
"""
Created on Mon May 28 08:43:33 2018

@author: 12600771
"""

import openpyxl
import xml.etree.ElementTree as ET

def convertToDevice():
    
    def indent(elem, level=0):
        i = "\n" + level*"  "
        j = "\n" + (level-1)*"  "
        if len(elem):
            if not elem.text or not elem.text.strip():
                elem.text = i + "  "
            if not elem.tail or not elem.tail.strip():
                elem.tail = i
            for subelem in elem:
                indent(subelem, level+1)
            if not elem.tail or not elem.tail.strip():
                elem.tail = j
        else:
            if level and (not elem.tail or not elem.tail.strip()):
                elem.tail = j
        return elem

    typeDict = {"bool":"Boolean", "double":"Double", "float": "Single", \
                "int":"Int32", "unsigned int":"UInt32"}

    fileName = "g_pChannelData.xlsx"
    workbook = openpyxl.load_workbook(fileName)
    sheet1 = workbook["Sheet1"]


    ns_xsi = "http://www.w3.org/2001/XMLSchema-instance"
    ns_xsd = "http://www.w3.org/2001/XMLSchema"
    device_attrib = {"xmlns:xsi":ns_xsi,"xmlns:xsd":ns_xsd}
    device = ET.Element("Device",attrib=device_attrib)
    # print(ET.tostring(root))


    # Create Protocol
    # <Device>
    #     <Protocol>
    protocol = ET.SubElement(device, 'Protocol', attrib={"xsi:type":"PVIProtocol"})

    # Create Variables
    # <Device>
    #   <Protocol>
    #       <Variables>
    variables = ET.SubElement(protocol, 'Variables')

    # Create ProtocolVariable
    # <Variables>
    #   <ProtocolVariable>
    #   <ProtocolVariable>
    for i in range(2, sheet1.max_row+1):
        protocolVariable = ET.SubElement(variables, 'ProtocolVariable', \
                                    attrib={"xsi:type":"PVIProtocolVariable"})

        # parse from excel to text
        num = 1
        cellForLongName = sheet1.cell(row = i, column = 4)
        longName = cellForLongName.value
        cellForType = sheet1.cell(row = i, column = 2)
        typeName = cellForType.value
        cellForShortName = sheet1.cell(row = i, column = 5)
        shortName = cellForShortName.value
        name = longName if len(longName) < 32 else shortName

        if "[" in longName:
            index1 = longName.find("[")
            index2 = longName.find("]")
            num = int(longName[index1+1:index2])

        if "[" in name:
            index1 = name.find("[")
            name = name[:index1]

        # text to xml
        """       
        <ProtocolVariable xsi:type="PVIProtocolVariable">
            <Name>P_DisconStateMachineStartTime</Name>
            <Description>Delay in discon state machine start up</Description>
            <VariableType>Single</VariableType>
            <ArraySize>1</ArraySize>
            <DefaultVal>10</DefaultVal>
            <MinVal />
            <MaxVal />
            <OutgoingScalingFactor />
            <Unit />
            <DisplayNames />
            <RemoteName>P_DisStaMacStaTim</RemoteName>
        </ProtocolVariable> 
        """

        # Do not parse array for the present
        if "[" not in name:
            ET.SubElement(protocolVariable, "Name").text = longName
            ET.SubElement(protocolVariable, "Description")
            ET.SubElement(protocolVariable, "VariableType")\
                                .text = typeDict.get(typeName)
            ET.SubElement(protocolVariable, "ArraySize").text = "1"
            ET.SubElement(protocolVariable, "DefaultVal")
            ET.SubElement(protocolVariable, "MinVal")
            ET.SubElement(protocolVariable, "MaxVal")
            ET.SubElement(protocolVariable, "OutgoingScalingFactor")
            ET.SubElement(protocolVariable, "Unit")
            ET.SubElement(protocolVariable, "DisplayNames")
            ET.SubElement(protocolVariable, "RemoteName").text = name

    # Add some other node
    """
    <Protocol>
        <Variables>
        </Variables>
        <RemoteIP>172.160.22.12</RemoteIP>
        <RemotePort>11169</RemotePort>
        <StructName>g_ChannelDataFull</StructName>
    """
    newStructTypeName = "g_ChannelDataFull"
    ET.SubElement(protocol, "RemoteIP").text = "172.160.22.12"
    ET.SubElement(protocol, "RemotePort").text = "11169"
    ET.SubElement(protocol, "StructName").text = newStructTypeName

    # <Device>
    #     <Name>Turbine Controller</Name>
    #     <Description>Symbols for the B&amp;R controller</Description>
    #     <ReadFrequencyHz>50</ReadFrequencyHz>
    #     <DataStorageConfig>
    #         <SourceFile>C:\Users\12600771\Desktop\device.device</SourceFile>
    #         <LastLoadedAt>2018-05-28T08:22:51.4819742+08:00</LastLoadedAt>
    #         <SourceFileLastModifiedAt>0001-01-01T00:00:00</SourceFileLastModifiedAt>
    #         <UpdateType>PromptForReload</UpdateType>
    #     </DataStorageConfig>

    ET.SubElement(device, "Name").text = "Turbine Controller"
    ET.SubElement(device, "Description").text \
        = "Symbols for the B&amp;R controller"
    ET.SubElement(device, "ReadFrequencyHz").text = str(50)
    dataStorageConfig  = ET.SubElement(device, "DataStorageConfig")
    ET.SubElement(dataStorageConfig, "SourceFile").text \
        = r"C:\Users\12600771\Desktop\device.device"
    ET.SubElement(dataStorageConfig, "LastLoadedAt").text\
        = "2018-05-28T08:22:51.4819742+08:00"
    ET.SubElement(dataStorageConfig, "LastLoadedAt").text \
        = "0001-01-01T00:00:00"
    ET.SubElement(dataStorageConfig, "UpdateType").text = "PromptForReload"

    # Add deviceConfig
    """ 
    <Device>
        <DeviceConfig>
            <Name>Turbine Controller</Name>
            <Assignments />
            <VariableConfig>
            </VariableConfig>
        </DeviceConfig>
    """
    deviceConfig = ET.SubElement(device, "DeviceConfig")
    ET.SubElement(deviceConfig, "Name").text = "Turbine Controller"
    ET.SubElement(deviceConfig, "Assignments")
    variableConfig = ET.SubElement(deviceConfig, "VariableConfig")

    # add protocolVariableConfig
    """ <VariableConfig>
        <ProtocolVariableConfig>
            <Name>CI_AlgSpeedControlSpeedSetpoint</Name>
            <Usage>ReadContinuously</Usage>
            <Record>true</Record>
        </ProtocolVariableConfig>
        ...
    </VariableConfig> """
    for i in range(2, sheet1.max_row+1):
        protocolVariableConfig = ET.SubElement(variableConfig, \
                                            "ProtocolVariableConfig")

        # parse from excel to text
        num = 1
        cellForLongName = sheet1.cell(row = i, column = 4)
        longName = cellForLongName.value
        cellForType = sheet1.cell(row = i, column = 2)
        typeName = cellForType.value
        cellForShortName = sheet1.cell(row = i, column = 5)
        shortName = cellForShortName.value
        name = longName if len(longName) < 32 else shortName

        if "[" not in name:
            ET.SubElement(protocolVariableConfig, "Name").text = longName
            ET.SubElement(protocolVariableConfig, "Usage").text \
                                        = "ReadContinuously"
            ET.SubElement(protocolVariableConfig, "Record").text = "true"

    # add more nodes
    """ <Device>
        <DependentDeviceNames />
        <Assignments />
        <Scripts>
            <ExecutableScript>
                <Expression>
                </Expression>
            </ExecutableScript>
        </Scripts>
    </Device> """
    ET.SubElement(device, "DependentDeviceNames")
    ET.SubElement(device, "Assignments")
    scripts = ET.SubElement(device, "Scripts")
    executableScript = ET.SubElement(scripts, "ExecutableScript")
    ET.SubElement(executableScript, "Expression")

    deviceFile = "device.device"
    indent(device)
    # ET.dump(device)
    tree = ET.ElementTree(device)
    tree.write(deviceFile, encoding="UTF-8", xml_declaration=True)
    # with open(deviceFile, 'w') as f:
    #     f.write(ET.tostring(device,encoding="UTF-8"))
    # ET.tostring(device)