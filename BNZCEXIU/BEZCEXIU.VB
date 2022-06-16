Imports System.Data
Imports NTSInformatica.CLN__STD
Imports System.Runtime.Remoting
Imports System.Runtime.Remoting.Channels
Imports System.Runtime.Remoting.Channels.Tcp
Imports NTSInformatica
Imports System.Globalization

Public Class CLEZCEXIU
    Inherits CLE__BASE

    Public oCldEXIU As CLDZCEXIU
    Public strNomeRamo As String = "BSZCEXIU"
    Public arrColonneIgnora As New List(Of String)(New String() {"ei_transmissiondate", "ei_issuedate"})


    Public path As String
    Public supplierId As String

#Region "Moduli"
    Private Moduli_P As Integer = CLN__STD.bsModAll
    Private ModuliExt_P As Integer = CLN__STD.bsModExtCGE
    Private ModuliSup_P As Integer = 0
    Private ModuliSupExt_P As Integer = 0
    Private ModuliPtn_P As Integer = 0
    Private ModuliPtnExt_P As Integer = 0

    Public ReadOnly Property Moduli() As Integer
        Get
            Return Moduli_P
        End Get
    End Property
    Public ReadOnly Property ModuliExt() As Integer
        Get
            Return ModuliExt_P
        End Get
    End Property
    Public ReadOnly Property ModuliSup() As Integer
        Get
            Return ModuliSup_P
        End Get
    End Property
    Public ReadOnly Property ModuliSupExt() As Integer
        Get
            Return ModuliSupExt_P
        End Get
    End Property
    Public ReadOnly Property ModuliPtn() As Integer
        Get
            Return ModuliPtn_P
        End Get
    End Property
    Public ReadOnly Property ModuliPtnExt() As Integer
        Get
            Return ModuliPtnExt_P
        End Get
    End Property
#End Region

    Public Overrides Function Init(ByRef App As CLE__APP,
                                    ByRef oScriptEngine As INT__SCRIPT, ByRef oCleLbmenu As Object, ByVal strTabella As String,
                                    ByVal bFiller1 As Boolean, ByVal strFiller1 As String,
                                    ByVal strFiller2 As String) As Boolean
        If MyBase.strNomeDal = "BD__BASE" Then MyBase.strNomeDal = "BDZCEXIU"
        MyBase.Init(App, oScriptEngine, oCleLbmenu, strTabella, False, "", "")
        oCldEXIU = CType(MyBase.ocldBase, CLDZCEXIU)
        oCldEXIU.Init(oApp)

        Return True
    End Function

    Public Overridable Function ApriImpostazioni(ByVal strDitta As String, ByRef ds As DataSet) As Boolean
        Dim dttImpostazioni As New DataTable
        Try
            If ds.Tables.Contains("IMPOSTAZIONI") Then Return True

            dttImpostazioni.TableName = "IMPOSTAZIONI"
            dttImpostazioni.Columns.Add("ei_username", GetType(String)).DefaultValue = ""
            dttImpostazioni.Columns.Add("ei_password", GetType(String)).DefaultValue = ""
            dttImpostazioni.Columns.Add("ei_fromto", GetType(String)).DefaultValue = "toM2"
            dttImpostazioni.Columns.Add("ei_legalname", GetType(String)).DefaultValue = ""
            dttImpostazioni.Columns.Add("ei_emaildest", GetType(String)).DefaultValue = ""
            dttImpostazioni.Columns.Add("ei_transmissiondate", GetType(Date)).DefaultValue = Now
            dttImpostazioni.Columns.Add("ei_partnerid", GetType(String)).DefaultValue = ""
            dttImpostazioni.Columns.Add("ei_issuedate", GetType(Date)).DefaultValue = Now
            dttImpostazioni.Columns.Add("ei_transmissionnumber", GetType(Integer)).DefaultValue = 0
            dttImpostazioni.Columns.Add("ei_auxhead1", GetType(String)).DefaultValue = "0009"
            dttImpostazioni.Columns.Add("ei_codlsar", GetType(Integer)).DefaultValue = 0
            dttImpostazioni.Columns.Add("xxei_codlsar", GetType(String)).DefaultValue = ""
            dttImpostazioni.Columns.Add("ei_whereart", GetType(String)).DefaultValue = ""
            dttImpostazioni.Columns.Add("ei_codfornlux", GetType(String)).DefaultValue = ""

            ds.Tables.Add(dttImpostazioni)
            dttImpostazioni.Rows.Add()

            Return True
        Catch ex As Exception
            '--------------------------------------------------------------

            CLN__STD.GestErr(ex, Me, "")

            '--------------------------------------------------------------
        End Try
        Return False
    End Function

    Public Overridable Function CaricaOpzReg() As Boolean

        path = (ocldBase.GetSettingBus("BSZCEXIU", "OPZIONI", ".", "path", "", " ", ""))
        supplierId = (ocldBase.GetSettingBus("BSZCEXIU", "OPZIONI", ".", "SupplierId", "", " ", ""))


        Return True

    End Function


    Public Overridable Function LeggiImpostazioni(ByVal strDitta As String, ByRef ds As DataSet) As Boolean
        Try
            ' copio da opzioni di registro i defaults
            For Each oCol As DataColumn In ds.Tables("IMPOSTAZIONI").Columns
                If arrColonneIgnora.Contains(oCol.ColumnName.ToLower) Then Continue For
                If oCol.ColumnName.ToLower.StartsWith("xx") Then Continue For
                ds.Tables("IMPOSTAZIONI").Rows(0)(oCol) = ocldBase.GetSettingBus(strNomeRamo, "OPZIONI", ".", "IMP_" & oCol.ColumnName, NTSCStr(oCol.DefaultValue), " ", NTSCStr(oCol.DefaultValue))
            Next

            Return True
        Catch ex As Exception
            '--------------------------------------------------------------

            CLN__STD.GestErr(ex, Me, "")

            '--------------------------------------------------------------
        End Try
        Return False
    End Function

    Public Overridable Function SalvaImpostazioni(ByVal strDitta As String, ByRef ds As DataSet) As Boolean
        Try
            ' copio da opzioni di registro i defaults
            For Each oCol As DataColumn In ds.Tables("IMPOSTAZIONI").Columns
                If arrColonneIgnora.Contains(oCol.ColumnName.ToLower) Then Continue For
                If oCol.ColumnName.ToLower.StartsWith("xx") Then Continue For
                ocldBase.SaveSettingBus(strNomeRamo, "OPZIONI", ".", "IMP_" & oCol.ColumnName, NTSCStr(ds.Tables("IMPOSTAZIONI").Rows(0)(oCol)), " ", False, False, False)
            Next

            Return True
        Catch ex As Exception
            '--------------------------------------------------------------

            CLN__STD.GestErr(ex, Me, "")

            '--------------------------------------------------------------
        End Try
        Return False
    End Function

    Public Overrides Function Apri(ByVal strDitta As String, ByRef ds As DataSet) As Boolean
        Dim dReturn As Boolean = False
        Try
            '--------------------------------------
            'tengo un puntatore al datatable che verrà usato nell'UI, quindi chiedo i dati al DAL
            strDittaCorrente = strDitta
            If ds Is Nothing Then ds = New DataSet
            If Not ApriImpostazioni(strDitta, ds) Then Return False

            '--------------------------------------------------------------
            'imposto i valori di default della tabella: con la riga che segue prendo prima i valori dal database
            SetDefaultValue(ds)

            dsShared = ds

            '--------------------------------------
            'creo gli eventi per la gestione del datatable dentro l'entity
            AddHandler dsShared.Tables("IMPOSTAZIONI").ColumnChanging, AddressOf BeforeColUpdate
            AddHandler dsShared.Tables("IMPOSTAZIONI").ColumnChanged, AddressOf AfterColUpdate

            If Not LeggiImpostazioni(strDitta, ds) Then Return False
            If Not oCldEXIU.GetArtprox(strDitta, ds) Then Return False

            bHasChanges = False

            Return True
        Catch ex As Exception
            '--------------------------------------------------------------

            CLN__STD.GestErr(ex, Me, "")

            '--------------------------------------------------------------
        End Try
    End Function

    Public Overrides Sub AfterColUpdate(ByVal sender As Object, ByVal e As DataColumnChangeEventArgs)
        Try
            'non valido la colonna se il dato non è cambiato
            If strPrevCelValue.IndexOf(e.Column.ColumnName.ToUpper + ";") > -1 Then
                strPrevCelValue = strPrevCelValue.Remove(strPrevCelValue.IndexOf(e.Column.ColumnName.ToUpper + ";"), e.Column.ColumnName.ToUpper.Length + 1)
                Return
            End If

            bHasChanges = True

            'comunico che una cella è cambiata, per fare in modo che se il dato è contenuto in una griglia 
            'vengano fatte le routine di validazione del caso
            ThrowRemoteEvent(New NTSEventArgs("GRIAGG", e.Column.Table.TableName & "§" & e.Column.ColumnName))

            e.Row.EndEdit()
            e.Row.EndEdit()

            '-------------------------------------------------------------
            'cerco e, se la trovo, eseguo la funzione specifica per la colonna modificata
            Dim strFunction As String = "AfterColUpdate_" & e.Column.ColumnName.ToLower
            Dim fun As System.Reflection.MethodInfo = Me.GetType.GetMethod(strFunction)  'occhio: è case_sensitive!!!!
            If Not fun Is Nothing Then fun.Invoke(Me, New Object() {sender, e})


        Catch ex As Exception
            '--------------------------------------------------------------

            CLN__STD.GestErr(ex, Me, "")

            '--------------------------------------------------------------
        End Try
    End Sub

    Public Overrides Sub BeforeColUpdate(ByVal sender As Object, ByVal e As DataColumnChangeEventArgs)
        Dim strErr As String = ""
        Try

            'memorizzo il valore corrente di cella per testarlo nella AfterColUpdate
            'solo se il dato è uguale a quello precedentemente contenuto nella cella
            If ValoriUguali(e.ProposedValue.ToString, e.Row(e.Column.ColumnName).ToString) Then
                strPrevCelValue += e.Column.ColumnName.ToUpper + ";"
                Return
            End If

            '-------------------------------------------------------------
            'controllo che in una cella short non venga inserito un numero troppo grande
            If Not CheckCellaShort(e, strErr) Then Throw New NTSException(oApp.Tr(Me, 128067822802778673, strErr))
            '-------------------------------------------------------------
            'cerco e, se la trovo, eseguo la funzione specifica per la colonna modificata
            Dim strFunction As String = "BeforeColUpdate_" & e.Column.ColumnName.ToLower
            Dim fun As System.Reflection.MethodInfo = Me.GetType.GetMethod(strFunction)  'occhio: è case_sensitive!!!!
            If Not fun Is Nothing Then fun.Invoke(Me, New Object() {sender, e})

        Catch ex As Exception
            '--------------------------------------------------------------

            CLN__STD.GestErr(ex, Me, "")

            '--------------------------------------------------------------
        End Try
    End Sub

    Public Overridable Sub BeforeColUpdate_ei_codlsar(ByVal sender As Object, ByVal e As DataColumnChangeEventArgs)
        Dim strDes As String = ""
        Dim dttTmp As New DataTable
        Dim bOk As Boolean = False
        Try
            If NTSCInt(e.ProposedValue) = 0 Then
                e.Row!xxei_codlsar = ""
            Else
                bOk = ocldBase.ValCodiceDb(e.ProposedValue.ToString, strDittaCorrente, "TABLSAR", "N", strDes, dttTmp)
                If Not bOk Then
                    e.ProposedValue = e.Row(e.Column)
                    ThrowRemoteEvent(New NTSEventArgs(CLN__STD.ThMsg.MSG_ERROR, "Il codice lista selezionata articoli indicato non esiste."))
                Else
                    e.Row!xxei_codlsar = strDes
                End If
            End If

        Catch ex As Exception
            '--------------------------------------------------------------

            CLN__STD.GestErr(ex, Me, "")

            '--------------------------------------------------------------
        End Try
    End Sub


    Public Overridable Sub AfterColUpdate_ei_codlsar(ByVal sender As Object, ByVal e As DataColumnChangeEventArgs)
        Try
            If Not oCldEXIU.GetArtprox(strDittaCorrente, e.Row.Table.DataSet) Then Return

        Catch ex As Exception
            '--------------------------------------------------------------

            CLN__STD.GestErr(ex, Me, "")

            '--------------------------------------------------------------
        End Try
    End Sub
    Public Overridable Sub AfterColUpdate_ei_whereart(ByVal sender As Object, ByVal e As DataColumnChangeEventArgs)
        Try
            If Not oCldEXIU.GetArtprox(strDittaCorrente, e.Row.Table.DataSet) Then Return

        Catch ex As Exception
            '--------------------------------------------------------------

            CLN__STD.GestErr(ex, Me, "")

            '--------------------------------------------------------------
        End Try
    End Sub


    Public Overridable Function Elabora(ByVal bModaleOutlook As Boolean) As Boolean
        Dim strNomeFile As String = ""
        Dim strPathFile As String = ""
        Dim strOggetto As String = ""
        Dim strMessaggio As String = ""
        Dim strContent As String = ""
        Dim strEmail As String = ""

        Dim riga As String = ""

        Dim sw As IO.StreamWriter





        Try
            CaricaOpzReg()

            'creo il file
            strNomeFile = "EXP_" & replace(cstr(Now.Date),"/","") & "_" & now.Hour & "_" & now.Minute & "_" & now.Second & ".TXT"


            sw = IO.File.CreateText(path & "\" & strNomeFile)

            riga = "SupplierProductId|SupplierId|StockQuantity|IncomingQuantity|IncomingDate|InvoiceCost|InvoiceNetCost"
            sw.WriteLine(riga)
            sw.Flush()


            ' per ognuno degli articoli
            For Each drArt As DataRow In dsShared.Tables("ARTPROX").Rows

                '   MsgBox(drArt!xx_qtadainv)

                'MsgBox(Format(drArt!xx_qtadainv, oApp.FormatQta))


                riga = (NTSCStr(drArt!ap_codart)) & "|" & NTSCStr(supplierId) & "|" & NTSCStr(Math.Round(NTSCDec(drArt!ap_esist), 3)) & "|" & Format(drArt!xx_qtadainv, oApp.FormatQta) & "|" & (NTSCDate(drArt!xx_dtinv).ToString("yyyyMMdd") & "|" & "|" & "|")
                sw.WriteLine(riga)
                sw.Flush()

            Next



            Return True
        Catch ex As Exception
            '--------------------------------------------------------------

            CLN__STD.GestErr(ex, Me, "")

            '--------------------------------------------------------------
        End Try
        Return False
    End Function

    Public Overridable Function InviaFile(ByVal bModaleOutlook As Boolean, ByVal strEmail As String, ByVal strOggetto As String, ByVal strMessaggio As String, ByVal strFile As String) As Boolean
        Dim oCleSend As CLE__SEND = Nothing
        Try
            '------------------------------------------------
            'creo e attivo l'entity e inizializzo la funzione che dovrà rilevare gli eventi dall'ENTITY
            Dim strErr As String = ""
            Dim oTmp As Object = Nothing
            If CLN__STD.NTSIstanziaDll(oApp.ServerDir, oApp.NetDir, "BNZCEXIU", "BE__SEND", oTmp, strErr, False, "", "") = False Then
                oApp.MsgBoxErr(oApp.Tr(Me, 128496233436616000, "ERRORE in fase di creazione Entity:" & vbCrLf & "|" & strErr & "|"))
                Return False
            End If
            oCleSend = CType(oTmp, CLE__SEND)
            If Not oCleSend.Init(oApp, oScript, oCleComm, "", False, "", "") Then
                oApp.MsgBoxErr(oApp.Tr(Me, 128496233436616000, "ERRORE in fase di init Entity:" & vbCrLf & "|" & strErr & "|"))
                Return False
            End If
            If Not oCleSend.InitExt(oApp) Then
                oApp.MsgBoxErr(oApp.Tr(Me, 128496233436616000, "ERRORE in fase di InitExt Entity:" & vbCrLf & "|" & strErr & "|"))
                Return False
            End If
            AddHandler oCleSend.RemoteEvent, AddressOf RedirectEventiEntity

            If Not oCleSend.InviaMail(strDittaCorrente, strEmail, "", "", strOggetto,
                                              New String() {strFile}, False,
                                              strMessaggio, strErr, strMessaggio,
                                              0, 0, bModaleOutlook,
                                              "BNZCEXIU") Then
                oApp.MsgBoxErr(oApp.Tr(Me, 128496233436616000, "ERRORE in fase di invio mail:" & vbCrLf & "|" & strErr & "|"))
                Return False
            End If

            'If Not oCleSend.InviaMail(strDittaCorrente, strEmail, strEmail, "", "Nuova password", New String() {}, False, strMessaggio, strErr) Then
            '    oApp.MsgBoxErr(oApp.Tr(Me, 128496233436616000, "ERRORE in fase di invio mail:" & vbCrLf & "|" & strErr & "|"))
            '    Return False
            'End If

            Return True
        Catch ex As Exception
            '--------------------------------------------------------------

            CLN__STD.GestErr(ex, Me, "")

            '--------------------------------------------------------------
        End Try
        Return False
    End Function

    Public Overridable Sub RedirectEventiEntity(sender As Object, ByRef e As NTSEventArgs)
        ThrowRemoteEvent(e)
    End Sub
End Class
