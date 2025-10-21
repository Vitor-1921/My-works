; =================================================================================
; === SCRIPT MESTRE DE AUTOMAÇÕES - VITOR (VERSÃO FINAL) ==========================
; =================================================================================
; Este arquivo combina todos os 6 scripts.
;
; RESUMO DOS ATALHOS:
; Ctrl + Q : Renomeia arquivo para o padrão [Inspection]
; Ctrl + W : Renomeia arquivo para o padrão [P&D]
; Ctrl + E : Renomeia arquivo para o padrão [EQF]
; Ctrl + D : Gera e cola e-mail de inspeção APROVADA
; Ctrl + G : Gera e cola e-mail de inspeção REPROVADA
; Ctrl + R : Cria uma pasta com o nome do texto copiado (Clipboard)
; =================================================================================

; --- Início da Seção de Renomear Arquivos ---

; Atalho: Ctrl + Q -> Renomeia para o padrão "Inspection"
^q::
{
    Clipboard := ""
    Send, ^c
    ClipWait, 1
    ArquivoAtual := Clipboard
    if (ArquivoAtual = "")
    {
        MsgBox, Falha ao copiar o caminho. Tente selecionar o arquivo novamente.
        return
    }
    if (FileExist(ArquivoAtual))
    {
        SplitPath, ArquivoAtual, NomeArquivo, DirPai, Extensao
        InfoExtra := "01 - [Inspection] - "
        NovoNome := InfoExtra . NomeArquivo . "." . Extensao
        NovoCaminho := DirPai . "\" . NovoNome
        FileMove, %ArquivoAtual%, %NovoCaminho%
        if ErrorLevel
        {
            MsgBox, Falha ao renomear o arquivo.
        }
    }
}
return

; Atalho: Ctrl + W -> Renomeia para o padrão "P&D"
^w::
{
    Clipboard := ""
    Send, ^c
    ClipWait, 1
    ArquivoAtual := Clipboard
    if (ArquivoAtual = "")
    {
        MsgBox, Falha ao copiar o caminho. Tente selecionar o arquivo novamente.
        return
    }
    if (FileExist(ArquivoAtual))
    {
        SplitPath, ArquivoAtual, NomeArquivo, DirPai, Extensao
        InfoExtra := "02 - [P&D] - "
        NovoNome := InfoExtra . NomeArquivo . "." . Extensao
        NovoCaminho := DirPai . "\" . NovoNome
        FileMove, %ArquivoAtual%, %NovoCaminho%
        if ErrorLevel
        {
            MsgBox, Falha ao renomear o arquivo.
        }
    }
}
return

; Atalho: Ctrl + E -> Renomeia para o padrão "EQF"
^e::
{
    Clipboard := ""
    Send, ^c
    ClipWait, 1
    ArquivoAtual := Clipboard
    if (ArquivoAtual = "")
    {
        MsgBox, Falha ao copiar o caminho. Tente selecionar o arquivo novamente.
        return
    }
    if (FileExist(ArquivoAtual))
    {
        SplitPath, ArquivoAtual, NomeArquivo, DirPai, Extensao
        InfoExtra := "03 - [EQF] - "
        NovoNome := InfoExtra . NomeArquivo . "." . Extensao
        NovoCaminho := DirPai . "\" . NovoNome
        FileMove, %ArquivoAtual%, %NovoCaminho%
        if ErrorLevel
        {
            MsgBox, Falha ao renomear o arquivo.
        }
    }
}
return

; --- Fim da Seção de Renomear Arquivos ---


; --- Início da Seção de E-mails Automáticos ---

; Atalho: Ctrl + D -> Gera e-mail de inspeção APROVADA
^d::
{
    InputBox, produto, Informar Produto, Qual é o produto?
    if (ErrorLevel)
        return
    InputBox, po, Informar PO, Qual é o PO?
    if (ErrorLevel)
        return
    mensagem := "The inspection of the product " . produto . " of the PO " . po . " is Approved.`n`n"
    mensagem .= "Please, check the shipment details with the logistics team.`n`n"
    mensagem .= "Regarding issues found during the inspection, please check our comments below:`n"
    mensagem .= "[Intelbras’ comments] It’s approved. The checklist will be updated."
    Clipboard := mensagem
    Sleep, 200
    Send, ^v
}
return

; Atalho: Ctrl + G -> Gera e-mail de inspeção REPROVADA
^g::
{
    InputBox, produto, Informar Produto, Qual é o produto?
    if (ErrorLevel)
        return
    InputBox, po, Informar PO, Qual é o PO?
    if (ErrorLevel)
        return
    mensagem := "The product " . produto . " of the PO " . po . " is Disapproved.`n`n"
    mensagem .= "Regarding issues found during the inspection, please check our comments below:`n`n"
    mensagem .= "[Intelbras’ comments] This is not acceptable.`n"
    mensagem .= "Now its necessary supplier’s team provide us an 8D report informing the root causes of these problems, what actions will be performed to solve the problem of this batch and what action will be performed to avoid these problems happens."
    Clipboard := mensagem
    Sleep, 200
    Send, ^v
}
return

; --- Fim da Seção de E-mails Automáticos ---


; --- Início da Seção de Criação de Pasta ---

; Atalho: Ctrl + R -> Cria pasta com o nome do conteúdo do Clipboard
^r::
{
    for window in ComObjCreate("Shell.Application").Windows
    {
        if (InStr(window.FullName, "explorer.exe"))
        {
            DirAtual := window.Document.Folder.Self.Path
            break
        }
    }

    if (DirAtual = "")
    {
        MsgBox, Falha ao capturar o caminho do diretório. Certifique-se de que uma janela do Explorer esteja ativa.
        return
    }

    NomePasta := Clipboard
    if (NomePasta = "")
    {
        MsgBox, O Clipboard está vazio. Copie um texto para usar como nome da pasta.
        return
    }

    ; Limpa caracteres inválidos para nomes de pasta
    NomePasta := RegExReplace(NomePasta, "[\r\n\t]", " ")
    NomePasta := RegExReplace(NomePasta, "[\\/:*?""<>|]", "")
    NomePasta := RegExReplace(NomePasta, "^\s+|\s+$") ; Remove espaços no início e no fim

    if (NomePasta = "")
    {
        MsgBox, O nome da pasta é inválido. Tente copiar outro texto.
        return
    }

    NovoCaminho := DirAtual . "\" . NomePasta

    if (FileExist(NovoCaminho))
    {
        MsgBox, A pasta "%NomePasta%" já existe neste local.
        return
    }

    FileCreateDir, %NovoCaminho%
    if ErrorLevel
    {
        MsgBox, Falha ao criar a pasta. Verifique as permissões.
    }
}
return

; --- Fim da Seção de Criação de Pasta ---
