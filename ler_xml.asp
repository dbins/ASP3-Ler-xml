<%
'Formato do XML de exemplo
'<xml>
'<campo1>x</campo1>
'<campo2>x</campo2>
'<campo3>x</campo3>
'<campo4>x</campo4>
'<campo5>
'	<campo51>x<campo51>
'	<campo52>x<campo52>
'	<campo53>x<campo53>
'</campo5>
'<campo5>
'	<campo51>x<campo51>
'	<campo52>x<campo52>
'	<campo53>x<campo53>
'</campo5>
'</xml>

Function ExcluirArquivo(ByRef Arquivo)
Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(Arquivo) = True Then
        'Excluindo o arquivo
		On Error resume Next
        objFSO.DeleteFile Arquivo, True
    End If
End Function

Function VerificarXML(ByRef NomeArquivo)
	set xmlDoc = Server.CreateObject("MSXML2.DOMDocument.3.0")
	xmlDoc.async = False
	xmlDoc.load(NomeArquivo) 'Aqui, carregamos os dados enviados
	
	If xmlDoc.parseError.errorCode <> 0 Then
		Mensagem = "Houve um erro ao ler o arquivo XML "
		Status = "ERRO"
	else
	
	'Lendo o conteudo do nó principal
	Set values_node = xmlDoc.getElementsByTagName("*")

	'Criando as variaveis principais - Troca normal
	Campo1 = XmlDoc.childNodes(1).childNodes(0).text
	Campo2 = XmlDoc.childNodes(1).childNodes(1).text
	Campo3 = XmlDoc.childNodes(1).childNodes(2).text
	Campo4 = XmlDoc.childNodes(1).childNodes(3).text

	Set values_node2 = xmlDoc.getElementsByTagName("campo5")
	For i = 0 to values_node2.length - 1
		Campo51 = values_node2.item(i).childNodes(0).text
		Campo52 = values_node2.item(i).childNodes(1).text
		Campo53 = values_node2.item(i).childNodes(2).text
	Next
	
	end if
	
	Set xmlDoc = Nothing
	Set values_node = Nothing
	ExcluirArquivo(NomeArquivo)

End Function

'Efetuando a leitura do XML enviado via request e a gravação num arquivo temporario
NomeArquivoXML=Month(Now) & Year(Now) & Day(Now) & Hour(Now) & Second(Now) & Minute(Now) & "arquivo.xml"

'Arquivo XML postado via POST numa variavel chamada TEXTAREA
'Gravando em disco primeiro antes de abrir
'Se quiser, pode guardar como log

Set fso = CreateObject("Scripting.FileSystemObject")
Set folderObject = fso.GetFolder("e:\xml\processados")
If Not NomeArquivoXML = "" Then
	On Error resume Next
    Set textStreamObject = folderObject.CreateTextFile(NomeArquivoXML, True, False)
    textStreamObject.WriteLine (Request("textarea"))
    Set textStreamObject = Nothing
End If

Set folderObject = Nothing
Set fso = Nothing

arquivo = "e:\xml\processados\" & NomeArquivoXML
VerificarXML(arquivo)
%>
