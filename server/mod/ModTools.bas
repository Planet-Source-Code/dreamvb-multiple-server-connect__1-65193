Attribute VB_Name = "ModTools"
Private Type ServCfg
    ServPort As Long
    ServMaxConn As Long
    ServName As String
End Type

Public TServCfg As ServCfg

Public Sub LoadCfg()
    TServCfg.ServName = GetSetting("VbServTest", "Cfg", "ServName", "~::Some Test Server::~")
    TServCfg.ServPort = Val(GetSetting("VbServTest", "Cfg", "Port", 90))
    TServCfg.ServMaxConn = Val(GetSetting("VbServTest", "Cfg", "MaxUsers", 20))
End Sub
