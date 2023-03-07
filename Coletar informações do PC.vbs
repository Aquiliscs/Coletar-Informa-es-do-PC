' Cria o objeto FileSystem para acessar o sistema de arquivos
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Cria o objeto WshShell para executar comandos no prompt de comando
Set objShell = WScript.CreateObject("WScript.Shell")

' Define o caminho do arquivo de texto na área de trabalho
strFilePath = objShell.SpecialFolders("Desktop") & "\informacoes_do_computador.txt"

' Cria o objeto TextStream para escrever no arquivo de texto
Set objFile = objFSO.CreateTextFile(strFilePath)

' Coleta e escreve as informações do processador
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")
For Each objItem in colItems
    objFile.WriteLine "Processador: " & objItem.Name
Next

' Coleta e escreve as informações da memória RAM
Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory")
total_memory = 0
For Each objItem in colItems
    objFile.WriteLine "Memória RAM: Total :" & objItem.Capacity / 1024 / 1024 & " MB Modelo :" & objItem.Manufacturer & " " & objItem.PartNumber
    total_memory = total_memory + objItem.Capacity / 1024 / 1024
Next


' Coleta e escreve as informações dos discos de armazenamento
Set colItems = objWMIService.ExecQuery("Select * from Win32_DiskDrive")
For Each objItem in colItems
    objFile.WriteLine "Disco: Modelo :" & objItem.Model & " Tamanho :" & FormatNumber(objItem.Size / 1024 / 1024 / 1024, 2) & " GB"
Next

' Coleta e escreve as informações das placas de rede
Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter")
For Each objItem in colItems
    If objItem.MACAddress <> "" Then
        objFile.WriteLine "Placa de Rede: " & objItem.Name & " MAC Address: " & objItem.MACAddress
    End If
Next

' Fecha o arquivo de texto
objFile.Close
