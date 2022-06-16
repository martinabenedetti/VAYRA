Imports System.Data
Imports NTSInformatica.CLN__STD

Public Class FRMZCEXIU
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
#Region "Variabili"
    Public oCleEXIU As CLEZCEXIU
    Public oCallParams As CLE__CLDP
    Public dsEXIU As DataSet
    Public dcEXIU As BindingSource = New BindingSource
    Public dcArtpro As New BindingSource
    Public lInstID As Integer = -1
    Public Const bsDbAzienda As Boolean = False
    Private components As System.ComponentModel.IContainer

    Public WithEvents barDockControlRight As DevExpress.XtraBars.BarDockControl
    Public WithEvents barDockControlLeft As DevExpress.XtraBars.BarDockControl
    Public WithEvents tlbMain As NTSInformatica.NTSBar
    Public WithEvents tlbStampa As NTSInformatica.NTSBarButtonItem
    Public WithEvents tlbStampaVideo As NTSInformatica.NTSBarButtonItem
    Public WithEvents tlbGuida As NTSInformatica.NTSBarButtonItem
    Public WithEvents tlbEsci As NTSInformatica.NTSBarButtonItem
    Public WithEvents barDockControlBottom As DevExpress.XtraBars.BarDockControl
    Public WithEvents barDockControlTop As DevExpress.XtraBars.BarDockControl
    Friend WithEvents tlbElabora As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents grpDatiAccount As NTSGroupBox
    Friend WithEvents lbEi_username As NTSLabel
    Friend WithEvents edEi_username As NTSTextBoxStr
    Friend WithEvents grArtpro As NTSGrid
    Friend WithEvents grvArtpro As NTSGridView
    Friend WithEvents grpDatiTrasmissione As NTSGroupBox
    Friend WithEvents lbEi_auxhead1 As NTSLabel
    Friend WithEvents edEi_auxhead1 As NTSComboBox
    Friend WithEvents lbEi_transmissionnumber As NTSLabel
    Friend WithEvents edEi_transmissionnumber As NTSTextBoxNum
    Friend WithEvents lbEi_issuedate As NTSLabel
    Friend WithEvents lbEi_partnerid As NTSLabel
    Friend WithEvents edEi_issuedate As NTSTextBoxData
    Friend WithEvents edEi_partnerid As NTSTextBoxStr
    Friend WithEvents grpDatiInvio As NTSGroupBox
    Friend WithEvents lbEi_emaildest As NTSLabel
    Friend WithEvents edEi_emaildest As NTSTextBoxStr
    Friend WithEvents lbEi_transmissiondate As NTSLabel
    Friend WithEvents edEi_transmissiondate As NTSTextBoxData
    Friend WithEvents cbEi_fromto As NTSComboBox
    Friend WithEvents lbEi_legalname As NTSLabel
    Friend WithEvents edEi_legalname As NTSTextBoxStr
    Friend WithEvents NtsLabel2 As NTSLabel
    Friend WithEvents lbEi_password As NTSLabel
    Friend WithEvents edEi_password As NTSTextBoxStr
    Friend WithEvents grpArtpro As NTSGroupBox
    Friend WithEvents lbEi_whereart As NTSLabel
    Friend WithEvents cmdScegliArticoli As NTSButton
    Friend WithEvents edEi_codlsar As NTSTextBoxNum
    Friend WithEvents lbXxEi_codlsar As NTSLabel
    Friend WithEvents lbEi_codlsar As NTSLabel
    Friend WithEvents tlbSalva As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents ap_codart As NTSGridColumn
    Friend WithEvents xxapx_codart As NTSGridColumn
    Friend WithEvents xxar_unmis As NTSGridColumn
    Friend WithEvents ap_esist As NTSGridColumn
    Friend WithEvents lbEi_codfornlux As NTSLabel
    Friend WithEvents edEi_codfornlux As NTSTextBoxStr
    Friend WithEvents ckModaleOutlook As NTSCheckBox
    Friend WithEvents xx_qtadainv As NTSGridColumn
    Friend WithEvents xx_dtinv As NTSGridColumn
    Public WithEvents NtsBarManager1 As NTSInformatica.NTSBarManager
#End Region

#Region "EVENTI DI FORM"
    Public Overridable Sub FRMZCEXIU_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            '-------------------------------------------------
            'predispongo i controlli
            InitControls()

            '-------------------------------------------------
            'leggo dal database i dati e collego il NTSBinding
            If Not oCleEXIU.Apri(DittaCorrente, dsEXIU) Then Me.Close()

            '-------------------------------------------------
            'collego il BindingSource ai vari controlli 
            Bindcontrols()

            '-------------------------------------------------
            'sempre alla fine di questa funzione: applico le regole della gctl
            'e comunque dopo si può lanciare dopo aver impostato la ditta, cambiato il tipodocumento, ecc ...
            'GctlTipoDoc = ""
            GctlSetRoules()

            ' Acquisisce InstID
            'lInstID = oMenu.GetTblInstId("hh_zcal_ttnrpag", False)

        Catch ex As Exception
            '-------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '-------------------------------------------------
        End Try
    End Sub

    Public Overridable Sub FRMZCEXIU_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Try
            ' Cancella i dati contenuti nella tabella temporanea (che hanno stesso InstID)
            'oMenu.ResetTblInstId("hh_zcal_ttnrpag", False, lInstID)

            dcEXIU.Dispose()
            dsEXIU.Dispose()
        Catch
        End Try
    End Sub
#End Region

    Public Overloads Function Init(ByRef Menu As CLE__MENU, ByRef Param As CLE__CLDP, Optional ByVal Ditta As String = "", Optional ByRef SharedControls As CLE__EVNT = Nothing) As Boolean
        oMenu = Menu
        oApp = oMenu.App
        oCallParams = Param
        If Ditta <> "" Then
            DittaCorrente = Ditta
        Else
            DittaCorrente = oApp.Ditta
        End If
        Me.GctlTipoDoc = ""

        InitializeComponent()
        Me.MinimumSize = Me.Size

        '------------------------------------------------
        'creo e attivo l'entity e inizializzo la funzione che dovrà rilevare gli eventi dall'ENTITY
        Dim strErr As String = ""
        Dim oTmp As Object = Nothing
        If CLN__STD.NTSIstanziaDll(oApp.ServerDir, oApp.NetDir, "BNZCEXIU", "BEZCEXIU", oTmp, strErr, False, "", "") = False Then
            oApp.MsgBoxErr(oApp.Tr(Me, 128589409211093750, "ERRORE in fase di creazione Entity:" & vbCrLf & "|" & strErr & "|"))
            Return False
        End If
        oCleEXIU = CType(oTmp, CLEZCEXIU)
        '------------------------------------------------

        AddHandler oCleEXIU.RemoteEvent, AddressOf GestisciEventiEntity
        If oCleEXIU.Init(oApp, NTSScript, oMenu.oCleComm, "", False, "", "") = False Then Return False

        Return True
    End Function

    Public Overridable Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FRMZCEXIU))
        Me.NtsBarManager1 = New NTSInformatica.NTSBarManager()
        Me.tlbMain = New NTSInformatica.NTSBar()
        Me.tlbSalva = New DevExpress.XtraBars.BarButtonItem()
        Me.tlbElabora = New DevExpress.XtraBars.BarButtonItem()
        Me.tlbStampa = New NTSInformatica.NTSBarButtonItem()
        Me.tlbStampaVideo = New NTSInformatica.NTSBarButtonItem()
        Me.tlbGuida = New NTSInformatica.NTSBarButtonItem()
        Me.tlbEsci = New NTSInformatica.NTSBarButtonItem()
        Me.barDockControlTop = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlBottom = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlLeft = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlRight = New DevExpress.XtraBars.BarDockControl()
        Me.grpDatiAccount = New NTSInformatica.NTSGroupBox()
        Me.lbEi_password = New NTSInformatica.NTSLabel()
        Me.edEi_password = New NTSInformatica.NTSTextBoxStr()
        Me.lbEi_username = New NTSInformatica.NTSLabel()
        Me.edEi_username = New NTSInformatica.NTSTextBoxStr()
        Me.grpDatiInvio = New NTSInformatica.NTSGroupBox()
        Me.ckModaleOutlook = New NTSInformatica.NTSCheckBox()
        Me.lbEi_emaildest = New NTSInformatica.NTSLabel()
        Me.edEi_emaildest = New NTSInformatica.NTSTextBoxStr()
        Me.lbEi_transmissiondate = New NTSInformatica.NTSLabel()
        Me.edEi_transmissiondate = New NTSInformatica.NTSTextBoxData()
        Me.cbEi_fromto = New NTSInformatica.NTSComboBox()
        Me.lbEi_legalname = New NTSInformatica.NTSLabel()
        Me.edEi_legalname = New NTSInformatica.NTSTextBoxStr()
        Me.NtsLabel2 = New NTSInformatica.NTSLabel()
        Me.grpDatiTrasmissione = New NTSInformatica.NTSGroupBox()
        Me.lbEi_codfornlux = New NTSInformatica.NTSLabel()
        Me.edEi_codfornlux = New NTSInformatica.NTSTextBoxStr()
        Me.lbEi_auxhead1 = New NTSInformatica.NTSLabel()
        Me.edEi_auxhead1 = New NTSInformatica.NTSComboBox()
        Me.lbEi_transmissionnumber = New NTSInformatica.NTSLabel()
        Me.edEi_transmissionnumber = New NTSInformatica.NTSTextBoxNum()
        Me.lbEi_issuedate = New NTSInformatica.NTSLabel()
        Me.lbEi_partnerid = New NTSInformatica.NTSLabel()
        Me.edEi_issuedate = New NTSInformatica.NTSTextBoxData()
        Me.edEi_partnerid = New NTSInformatica.NTSTextBoxStr()
        Me.grArtpro = New NTSInformatica.NTSGrid()
        Me.grvArtpro = New NTSInformatica.NTSGridView()
        Me.ap_codart = New NTSInformatica.NTSGridColumn()
        Me.xxapx_codart = New NTSInformatica.NTSGridColumn()
        Me.xxar_unmis = New NTSInformatica.NTSGridColumn()
        Me.ap_esist = New NTSInformatica.NTSGridColumn()
        Me.xx_qtadainv = New NTSInformatica.NTSGridColumn()
        Me.xx_dtinv = New NTSInformatica.NTSGridColumn()
        Me.grpArtpro = New NTSInformatica.NTSGroupBox()
        Me.edEi_codlsar = New NTSInformatica.NTSTextBoxNum()
        Me.lbXxEi_codlsar = New NTSInformatica.NTSLabel()
        Me.lbEi_codlsar = New NTSInformatica.NTSLabel()
        Me.lbEi_whereart = New NTSInformatica.NTSLabel()
        Me.cmdScegliArticoli = New NTSInformatica.NTSButton()
        CType(Me.dttSmartArt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NtsBarManager1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpDatiAccount, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpDatiAccount.SuspendLayout()
        CType(Me.edEi_password.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.edEi_username.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpDatiInvio, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpDatiInvio.SuspendLayout()
        CType(Me.ckModaleOutlook.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.edEi_emaildest.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.edEi_transmissiondate.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cbEi_fromto.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.edEi_legalname.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpDatiTrasmissione, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpDatiTrasmissione.SuspendLayout()
        CType(Me.edEi_codfornlux.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.edEi_auxhead1.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.edEi_transmissionnumber.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.edEi_issuedate.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.edEi_partnerid.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grArtpro, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grvArtpro, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpArtpro, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpArtpro.SuspendLayout()
        CType(Me.edEi_codlsar.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'NtsBarManager1
        '
        Me.NtsBarManager1.AllowCustomization = False
        Me.NtsBarManager1.Bars.AddRange(New DevExpress.XtraBars.Bar() {Me.tlbMain})
        Me.NtsBarManager1.DockControls.Add(Me.barDockControlTop)
        Me.NtsBarManager1.DockControls.Add(Me.barDockControlBottom)
        Me.NtsBarManager1.DockControls.Add(Me.barDockControlLeft)
        Me.NtsBarManager1.DockControls.Add(Me.barDockControlRight)
        Me.NtsBarManager1.Form = Me
        Me.NtsBarManager1.Items.AddRange(New DevExpress.XtraBars.BarItem() {Me.tlbStampa, Me.tlbStampaVideo, Me.tlbGuida, Me.tlbEsci, Me.tlbElabora, Me.tlbSalva})
        Me.NtsBarManager1.MaxItemId = 6
        '
        'tlbMain
        '
        Me.tlbMain.BarName = "tlbMain"
        Me.tlbMain.DockCol = 0
        Me.tlbMain.DockRow = 0
        Me.tlbMain.DockStyle = DevExpress.XtraBars.BarDockStyle.Top
        Me.tlbMain.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {New DevExpress.XtraBars.LinkPersistInfo(Me.tlbSalva), New DevExpress.XtraBars.LinkPersistInfo(Me.tlbElabora), New DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, False, Me.tlbStampa, True), New DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, False, Me.tlbStampaVideo, False), New DevExpress.XtraBars.LinkPersistInfo(Me.tlbGuida, True), New DevExpress.XtraBars.LinkPersistInfo(Me.tlbEsci)})
        Me.tlbMain.OptionsBar.AllowQuickCustomization = False
        Me.tlbMain.OptionsBar.DisableClose = True
        Me.tlbMain.OptionsBar.DrawDragBorder = False
        Me.tlbMain.OptionsBar.UseWholeRow = True
        Me.tlbMain.Text = "tlbMain"
        '
        'tlbSalva
        '
        Me.tlbSalva.Caption = "Salva"
        Me.tlbSalva.Glyph = CType(resources.GetObject("tlbSalva.Glyph"), System.Drawing.Image)
        Me.tlbSalva.Id = 5
        Me.tlbSalva.Name = "tlbSalva"
        '
        'tlbElabora
        '
        Me.tlbElabora.Caption = "Elabora"
        Me.tlbElabora.Glyph = CType(resources.GetObject("tlbElabora.Glyph"), System.Drawing.Image)
        Me.tlbElabora.Id = 4
        Me.tlbElabora.Name = "tlbElabora"
        '
        'tlbStampa
        '
        Me.tlbStampa.Caption = "Stampa"
        Me.tlbStampa.Enabled = False
        Me.tlbStampa.Glyph = CType(resources.GetObject("tlbStampa.Glyph"), System.Drawing.Image)
        Me.tlbStampa.Id = 0
        Me.tlbStampa.Name = "tlbStampa"
        Me.tlbStampa.Visible = False
        '
        'tlbStampaVideo
        '
        Me.tlbStampaVideo.Caption = "StampaVideo"
        Me.tlbStampaVideo.Enabled = False
        Me.tlbStampaVideo.Glyph = CType(resources.GetObject("tlbStampaVideo.Glyph"), System.Drawing.Image)
        Me.tlbStampaVideo.Id = 1
        Me.tlbStampaVideo.Name = "tlbStampaVideo"
        Me.tlbStampaVideo.Visible = False
        '
        'tlbGuida
        '
        Me.tlbGuida.Caption = "Guida"
        Me.tlbGuida.Glyph = CType(resources.GetObject("tlbGuida.Glyph"), System.Drawing.Image)
        Me.tlbGuida.Id = 2
        Me.tlbGuida.Name = "tlbGuida"
        Me.tlbGuida.Visible = True
        '
        'tlbEsci
        '
        Me.tlbEsci.Caption = "Esci"
        Me.tlbEsci.Glyph = CType(resources.GetObject("tlbEsci.Glyph"), System.Drawing.Image)
        Me.tlbEsci.Id = 3
        Me.tlbEsci.Name = "tlbEsci"
        Me.tlbEsci.Visible = True
        '
        'barDockControlTop
        '
        Me.barDockControlTop.CausesValidation = False
        Me.barDockControlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.barDockControlTop.Location = New System.Drawing.Point(0, 0)
        Me.barDockControlTop.Size = New System.Drawing.Size(640, 35)
        '
        'barDockControlBottom
        '
        Me.barDockControlBottom.CausesValidation = False
        Me.barDockControlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.barDockControlBottom.Location = New System.Drawing.Point(0, 511)
        Me.barDockControlBottom.Size = New System.Drawing.Size(640, 0)
        '
        'barDockControlLeft
        '
        Me.barDockControlLeft.CausesValidation = False
        Me.barDockControlLeft.Dock = System.Windows.Forms.DockStyle.Left
        Me.barDockControlLeft.Location = New System.Drawing.Point(0, 35)
        Me.barDockControlLeft.Size = New System.Drawing.Size(0, 476)
        '
        'barDockControlRight
        '
        Me.barDockControlRight.CausesValidation = False
        Me.barDockControlRight.Dock = System.Windows.Forms.DockStyle.Right
        Me.barDockControlRight.Location = New System.Drawing.Point(640, 35)
        Me.barDockControlRight.Size = New System.Drawing.Size(0, 476)
        '
        'grpDatiAccount
        '
        Me.grpDatiAccount.AllowDrop = True
        Me.grpDatiAccount.Appearance.BackColor = System.Drawing.Color.Transparent
        Me.grpDatiAccount.Appearance.Options.UseBackColor = True
        Me.grpDatiAccount.Controls.Add(Me.lbEi_password)
        Me.grpDatiAccount.Controls.Add(Me.edEi_password)
        Me.grpDatiAccount.Controls.Add(Me.lbEi_username)
        Me.grpDatiAccount.Controls.Add(Me.edEi_username)
        Me.grpDatiAccount.Location = New System.Drawing.Point(624, 94)
        Me.grpDatiAccount.Name = "grpDatiAccount"
        Me.grpDatiAccount.Size = New System.Drawing.Size(16, 64)
        Me.grpDatiAccount.Text = "DATI ACCOUNT"
        Me.grpDatiAccount.Visible = False
        '
        'lbEi_password
        '
        Me.lbEi_password.BackColor = System.Drawing.Color.Transparent
        Me.lbEi_password.Location = New System.Drawing.Point(4, 48)
        Me.lbEi_password.Name = "lbEi_password"
        Me.lbEi_password.NTSBordeStyle = NTSInformatica.NTSLabel.NTSBorderStyle.FieldCaption
        Me.lbEi_password.Size = New System.Drawing.Size(72, 20)
        Me.lbEi_password.Text = "Password"
        Me.lbEi_password.UseMnemonic = False
        '
        'edEi_password
        '
        Me.edEi_password.Location = New System.Drawing.Point(80, 48)
        Me.edEi_password.Name = "edEi_password"
        Me.edEi_password.Properties.AppearanceDisabled.ForeColor = System.Drawing.Color.Black
        Me.edEi_password.Properties.AppearanceDisabled.Options.UseForeColor = True
        Me.edEi_password.Properties.AutoHeight = False
        Me.edEi_password.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.edEi_password.Properties.MaxLength = 65536
        Me.edEi_password.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.edEi_password.Size = New System.Drawing.Size(152, 20)
        '
        'lbEi_username
        '
        Me.lbEi_username.BackColor = System.Drawing.Color.Transparent
        Me.lbEi_username.Location = New System.Drawing.Point(4, 24)
        Me.lbEi_username.Name = "lbEi_username"
        Me.lbEi_username.NTSBordeStyle = NTSInformatica.NTSLabel.NTSBorderStyle.FieldCaption
        Me.lbEi_username.Size = New System.Drawing.Size(72, 20)
        Me.lbEi_username.Text = "Username"
        Me.lbEi_username.UseMnemonic = False
        '
        'edEi_username
        '
        Me.edEi_username.Location = New System.Drawing.Point(80, 24)
        Me.edEi_username.Name = "edEi_username"
        Me.edEi_username.Properties.AppearanceDisabled.ForeColor = System.Drawing.Color.Black
        Me.edEi_username.Properties.AppearanceDisabled.Options.UseForeColor = True
        Me.edEi_username.Properties.AutoHeight = False
        Me.edEi_username.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.edEi_username.Properties.MaxLength = 65536
        Me.edEi_username.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.edEi_username.Size = New System.Drawing.Size(152, 20)
        '
        'grpDatiInvio
        '
        Me.grpDatiInvio.AllowDrop = True
        Me.grpDatiInvio.Appearance.BackColor = System.Drawing.Color.Transparent
        Me.grpDatiInvio.Appearance.Options.UseBackColor = True
        Me.grpDatiInvio.Controls.Add(Me.ckModaleOutlook)
        Me.grpDatiInvio.Controls.Add(Me.lbEi_emaildest)
        Me.grpDatiInvio.Controls.Add(Me.edEi_emaildest)
        Me.grpDatiInvio.Controls.Add(Me.lbEi_transmissiondate)
        Me.grpDatiInvio.Controls.Add(Me.edEi_transmissiondate)
        Me.grpDatiInvio.Controls.Add(Me.cbEi_fromto)
        Me.grpDatiInvio.Controls.Add(Me.lbEi_legalname)
        Me.grpDatiInvio.Controls.Add(Me.edEi_legalname)
        Me.grpDatiInvio.Controls.Add(Me.NtsLabel2)
        Me.grpDatiInvio.Location = New System.Drawing.Point(620, 42)
        Me.grpDatiInvio.Name = "grpDatiInvio"
        Me.grpDatiInvio.Size = New System.Drawing.Size(20, 32)
        Me.grpDatiInvio.Text = "DATI INVIO"
        Me.grpDatiInvio.Visible = False
        '
        'ckModaleOutlook
        '
        Me.ckModaleOutlook.Location = New System.Drawing.Point(4, 120)
        Me.ckModaleOutlook.Name = "ckModaleOutlook"
        Me.ckModaleOutlook.Properties.Appearance.BackColor = System.Drawing.Color.Transparent
        Me.ckModaleOutlook.Properties.Appearance.Options.UseBackColor = True
        Me.ckModaleOutlook.Properties.AutoHeight = False
        Me.ckModaleOutlook.Properties.Caption = "Visualizza prima di inviare"
        Me.ckModaleOutlook.Properties.CheckStyle = DevExpress.XtraEditors.Controls.CheckStyles.UserDefined
        Me.ckModaleOutlook.Properties.PictureChecked = CType(resources.GetObject("ckModaleOutlook.Properties.PictureChecked"), System.Drawing.Image)
        Me.ckModaleOutlook.Properties.PictureUnchecked = CType(resources.GetObject("ckModaleOutlook.Properties.PictureUnchecked"), System.Drawing.Image)
        Me.ckModaleOutlook.Size = New System.Drawing.Size(228, 20)
        '
        'lbEi_emaildest
        '
        Me.lbEi_emaildest.BackColor = System.Drawing.Color.Transparent
        Me.lbEi_emaildest.Location = New System.Drawing.Point(4, 72)
        Me.lbEi_emaildest.Name = "lbEi_emaildest"
        Me.lbEi_emaildest.NTSBordeStyle = NTSInformatica.NTSLabel.NTSBorderStyle.FieldCaption
        Me.lbEi_emaildest.Size = New System.Drawing.Size(72, 20)
        Me.lbEi_emaildest.Text = "Email dest."
        Me.lbEi_emaildest.UseMnemonic = False
        '
        'edEi_emaildest
        '
        Me.edEi_emaildest.Location = New System.Drawing.Point(80, 72)
        Me.edEi_emaildest.Name = "edEi_emaildest"
        Me.edEi_emaildest.Properties.AppearanceDisabled.ForeColor = System.Drawing.Color.Black
        Me.edEi_emaildest.Properties.AppearanceDisabled.Options.UseForeColor = True
        Me.edEi_emaildest.Properties.AutoHeight = False
        Me.edEi_emaildest.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.edEi_emaildest.Properties.MaxLength = 65536
        Me.edEi_emaildest.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.edEi_emaildest.Size = New System.Drawing.Size(152, 20)
        '
        'lbEi_transmissiondate
        '
        Me.lbEi_transmissiondate.BackColor = System.Drawing.Color.Transparent
        Me.lbEi_transmissiondate.Location = New System.Drawing.Point(4, 96)
        Me.lbEi_transmissiondate.Name = "lbEi_transmissiondate"
        Me.lbEi_transmissiondate.NTSBordeStyle = NTSInformatica.NTSLabel.NTSBorderStyle.FieldCaption
        Me.lbEi_transmissiondate.Size = New System.Drawing.Size(72, 20)
        Me.lbEi_transmissiondate.Text = "Trasmesso il"
        Me.lbEi_transmissiondate.UseMnemonic = False
        '
        'edEi_transmissiondate
        '
        Me.edEi_transmissiondate.Location = New System.Drawing.Point(80, 96)
        Me.edEi_transmissiondate.Name = "edEi_transmissiondate"
        Me.edEi_transmissiondate.Properties.AppearanceDisabled.ForeColor = System.Drawing.Color.Black
        Me.edEi_transmissiondate.Properties.AppearanceDisabled.Options.UseForeColor = True
        Me.edEi_transmissiondate.Properties.AutoHeight = False
        Me.edEi_transmissiondate.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.edEi_transmissiondate.Properties.MaxLength = 65536
        Me.edEi_transmissiondate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.edEi_transmissiondate.Size = New System.Drawing.Size(152, 20)
        '
        'cbEi_fromto
        '
        Me.cbEi_fromto.Location = New System.Drawing.Point(80, 24)
        Me.cbEi_fromto.Name = "cbEi_fromto"
        Me.cbEi_fromto.Properties.AutoHeight = False
        Me.cbEi_fromto.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cbEi_fromto.Properties.DropDownRows = 30
        Me.cbEi_fromto.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        Me.cbEi_fromto.Size = New System.Drawing.Size(152, 20)
        '
        'lbEi_legalname
        '
        Me.lbEi_legalname.BackColor = System.Drawing.Color.Transparent
        Me.lbEi_legalname.Location = New System.Drawing.Point(4, 48)
        Me.lbEi_legalname.Name = "lbEi_legalname"
        Me.lbEi_legalname.NTSBordeStyle = NTSInformatica.NTSLabel.NTSBorderStyle.FieldCaption
        Me.lbEi_legalname.Size = New System.Drawing.Size(72, 20)
        Me.lbEi_legalname.Text = "Azienda"
        Me.lbEi_legalname.UseMnemonic = False
        '
        'edEi_legalname
        '
        Me.edEi_legalname.Location = New System.Drawing.Point(80, 48)
        Me.edEi_legalname.Name = "edEi_legalname"
        Me.edEi_legalname.Properties.AppearanceDisabled.ForeColor = System.Drawing.Color.Black
        Me.edEi_legalname.Properties.AppearanceDisabled.Options.UseForeColor = True
        Me.edEi_legalname.Properties.AutoHeight = False
        Me.edEi_legalname.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.edEi_legalname.Properties.MaxLength = 65536
        Me.edEi_legalname.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.edEi_legalname.Size = New System.Drawing.Size(152, 20)
        '
        'NtsLabel2
        '
        Me.NtsLabel2.BackColor = System.Drawing.Color.Transparent
        Me.NtsLabel2.Location = New System.Drawing.Point(4, 24)
        Me.NtsLabel2.Name = "NtsLabel2"
        Me.NtsLabel2.NTSBordeStyle = NTSInformatica.NTSLabel.NTSBorderStyle.FieldCaption
        Me.NtsLabel2.Size = New System.Drawing.Size(72, 20)
        Me.NtsLabel2.Text = "Direzione"
        Me.NtsLabel2.UseMnemonic = False
        '
        'grpDatiTrasmissione
        '
        Me.grpDatiTrasmissione.AllowDrop = True
        Me.grpDatiTrasmissione.Appearance.BackColor = System.Drawing.Color.Transparent
        Me.grpDatiTrasmissione.Appearance.Options.UseBackColor = True
        Me.grpDatiTrasmissione.Controls.Add(Me.lbEi_codfornlux)
        Me.grpDatiTrasmissione.Controls.Add(Me.edEi_codfornlux)
        Me.grpDatiTrasmissione.Controls.Add(Me.lbEi_auxhead1)
        Me.grpDatiTrasmissione.Controls.Add(Me.edEi_auxhead1)
        Me.grpDatiTrasmissione.Controls.Add(Me.lbEi_transmissionnumber)
        Me.grpDatiTrasmissione.Controls.Add(Me.edEi_transmissionnumber)
        Me.grpDatiTrasmissione.Controls.Add(Me.lbEi_issuedate)
        Me.grpDatiTrasmissione.Controls.Add(Me.lbEi_partnerid)
        Me.grpDatiTrasmissione.Controls.Add(Me.edEi_issuedate)
        Me.grpDatiTrasmissione.Controls.Add(Me.edEi_partnerid)
        Me.grpDatiTrasmissione.Location = New System.Drawing.Point(624, 166)
        Me.grpDatiTrasmissione.Name = "grpDatiTrasmissione"
        Me.grpDatiTrasmissione.Size = New System.Drawing.Size(16, 72)
        Me.grpDatiTrasmissione.Text = "DATI TRASMISSIONE"
        Me.grpDatiTrasmissione.Visible = False
        '
        'lbEi_codfornlux
        '
        Me.lbEi_codfornlux.BackColor = System.Drawing.Color.Transparent
        Me.lbEi_codfornlux.Location = New System.Drawing.Point(4, 48)
        Me.lbEi_codfornlux.Name = "lbEi_codfornlux"
        Me.lbEi_codfornlux.NTSBordeStyle = NTSInformatica.NTSLabel.NTSBorderStyle.FieldCaption
        Me.lbEi_codfornlux.Size = New System.Drawing.Size(80, 20)
        Me.lbEi_codfornlux.Text = "cod. forn. lux"
        Me.lbEi_codfornlux.UseMnemonic = False
        '
        'edEi_codfornlux
        '
        Me.edEi_codfornlux.Location = New System.Drawing.Point(80, 48)
        Me.edEi_codfornlux.Name = "edEi_codfornlux"
        Me.edEi_codfornlux.Properties.AppearanceDisabled.ForeColor = System.Drawing.Color.Black
        Me.edEi_codfornlux.Properties.AppearanceDisabled.Options.UseForeColor = True
        Me.edEi_codfornlux.Properties.AutoHeight = False
        Me.edEi_codfornlux.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.edEi_codfornlux.Properties.MaxLength = 65536
        Me.edEi_codfornlux.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.edEi_codfornlux.Size = New System.Drawing.Size(152, 20)
        '
        'lbEi_auxhead1
        '
        Me.lbEi_auxhead1.BackColor = System.Drawing.Color.Transparent
        Me.lbEi_auxhead1.Location = New System.Drawing.Point(4, 120)
        Me.lbEi_auxhead1.Name = "lbEi_auxhead1"
        Me.lbEi_auxhead1.NTSBordeStyle = NTSInformatica.NTSLabel.NTSBorderStyle.FieldCaption
        Me.lbEi_auxhead1.Size = New System.Drawing.Size(72, 20)
        Me.lbEi_auxhead1.Text = "Riservati a"
        Me.lbEi_auxhead1.UseMnemonic = False
        '
        'edEi_auxhead1
        '
        Me.edEi_auxhead1.Location = New System.Drawing.Point(80, 120)
        Me.edEi_auxhead1.Name = "edEi_auxhead1"
        Me.edEi_auxhead1.Properties.AutoHeight = False
        Me.edEi_auxhead1.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.edEi_auxhead1.Properties.DropDownRows = 30
        Me.edEi_auxhead1.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        Me.edEi_auxhead1.Size = New System.Drawing.Size(152, 20)
        '
        'lbEi_transmissionnumber
        '
        Me.lbEi_transmissionnumber.BackColor = System.Drawing.Color.Transparent
        Me.lbEi_transmissionnumber.Location = New System.Drawing.Point(4, 96)
        Me.lbEi_transmissionnumber.Name = "lbEi_transmissionnumber"
        Me.lbEi_transmissionnumber.NTSBordeStyle = NTSInformatica.NTSLabel.NTSBorderStyle.FieldCaption
        Me.lbEi_transmissionnumber.Size = New System.Drawing.Size(72, 20)
        Me.lbEi_transmissionnumber.Text = "Progressivo"
        Me.lbEi_transmissionnumber.UseMnemonic = False
        '
        'edEi_transmissionnumber
        '
        Me.edEi_transmissionnumber.Location = New System.Drawing.Point(80, 96)
        Me.edEi_transmissionnumber.Name = "edEi_transmissionnumber"
        Me.edEi_transmissionnumber.Properties.Appearance.Options.UseTextOptions = True
        Me.edEi_transmissionnumber.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        Me.edEi_transmissionnumber.Properties.AppearanceDisabled.ForeColor = System.Drawing.Color.Black
        Me.edEi_transmissionnumber.Properties.AppearanceDisabled.Options.UseForeColor = True
        Me.edEi_transmissionnumber.Properties.AutoHeight = False
        Me.edEi_transmissionnumber.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.edEi_transmissionnumber.Properties.MaxLength = 65536
        Me.edEi_transmissionnumber.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.edEi_transmissionnumber.Size = New System.Drawing.Size(152, 20)
        '
        'lbEi_issuedate
        '
        Me.lbEi_issuedate.BackColor = System.Drawing.Color.Transparent
        Me.lbEi_issuedate.Location = New System.Drawing.Point(4, 72)
        Me.lbEi_issuedate.Name = "lbEi_issuedate"
        Me.lbEi_issuedate.NTSBordeStyle = NTSInformatica.NTSLabel.NTSBorderStyle.FieldCaption
        Me.lbEi_issuedate.Size = New System.Drawing.Size(72, 20)
        Me.lbEi_issuedate.Text = "Elaborato il"
        Me.lbEi_issuedate.UseMnemonic = False
        '
        'lbEi_partnerid
        '
        Me.lbEi_partnerid.BackColor = System.Drawing.Color.Transparent
        Me.lbEi_partnerid.Location = New System.Drawing.Point(4, 24)
        Me.lbEi_partnerid.Name = "lbEi_partnerid"
        Me.lbEi_partnerid.NTSBordeStyle = NTSInformatica.NTSLabel.NTSBorderStyle.FieldCaption
        Me.lbEi_partnerid.Size = New System.Drawing.Size(72, 20)
        Me.lbEi_partnerid.Text = "ID Partner"
        Me.lbEi_partnerid.UseMnemonic = False
        '
        'edEi_issuedate
        '
        Me.edEi_issuedate.Location = New System.Drawing.Point(80, 72)
        Me.edEi_issuedate.Name = "edEi_issuedate"
        Me.edEi_issuedate.Properties.AppearanceDisabled.ForeColor = System.Drawing.Color.Black
        Me.edEi_issuedate.Properties.AppearanceDisabled.Options.UseForeColor = True
        Me.edEi_issuedate.Properties.AutoHeight = False
        Me.edEi_issuedate.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.edEi_issuedate.Properties.MaxLength = 65536
        Me.edEi_issuedate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.edEi_issuedate.Size = New System.Drawing.Size(152, 20)
        '
        'edEi_partnerid
        '
        Me.edEi_partnerid.Location = New System.Drawing.Point(80, 24)
        Me.edEi_partnerid.Name = "edEi_partnerid"
        Me.edEi_partnerid.Properties.AppearanceDisabled.ForeColor = System.Drawing.Color.Black
        Me.edEi_partnerid.Properties.AppearanceDisabled.Options.UseForeColor = True
        Me.edEi_partnerid.Properties.AutoHeight = False
        Me.edEi_partnerid.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.edEi_partnerid.Properties.MaxLength = 65536
        Me.edEi_partnerid.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.edEi_partnerid.Size = New System.Drawing.Size(152, 20)
        '
        'grArtpro
        '
        Me.grArtpro.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grArtpro.Location = New System.Drawing.Point(4, 72)
        Me.grArtpro.MainView = Me.grvArtpro
        Me.grArtpro.Name = "grArtpro"
        Me.grArtpro.Size = New System.Drawing.Size(460, 96)
        Me.grArtpro.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.grvArtpro})
        '
        'grvArtpro
        '
        Me.grvArtpro.ActiveFilterEnabled = False
        Me.grvArtpro.Columns.AddRange(New DevExpress.XtraGrid.Columns.GridColumn() {Me.ap_codart, Me.xxapx_codart, Me.xxar_unmis, Me.ap_esist, Me.xx_qtadainv, Me.xx_dtinv})
        Me.grvArtpro.Enabled = True
        Me.grvArtpro.GridControl = Me.grArtpro
        Me.grvArtpro.GroupPanelText = "Per raggruppare: Tasto DX sulla colonna -> Group by this column *** Per totali pa" &
    "rziali/generali: sul piede di gruppo della colonna tasto DX -> SUM "
        Me.grvArtpro.Name = "grvArtpro"
        Me.grvArtpro.OptionsCustomization.AllowRowSizing = True
        Me.grvArtpro.OptionsFilter.AllowFilterEditor = False
        Me.grvArtpro.OptionsNavigation.EnterMoveNextColumn = True
        Me.grvArtpro.OptionsNavigation.UseTabKey = False
        Me.grvArtpro.OptionsSelection.EnableAppearanceFocusedRow = False
        Me.grvArtpro.OptionsView.ColumnAutoWidth = False
        Me.grvArtpro.OptionsView.GroupFooterShowMode = DevExpress.XtraGrid.Views.Grid.GroupFooterShowMode.Hidden
        Me.grvArtpro.OptionsView.ShowGroupPanel = False
        '
        'ap_codart
        '
        Me.ap_codart.AppearanceCell.Options.UseBackColor = True
        Me.ap_codart.AppearanceCell.Options.UseTextOptions = True
        Me.ap_codart.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
        Me.ap_codart.Caption = "Articolo"
        Me.ap_codart.Enabled = False
        Me.ap_codart.FieldName = "ap_codart"
        Me.ap_codart.FilterMode = DevExpress.XtraGrid.ColumnFilterMode.DisplayText
        Me.ap_codart.Name = "ap_codart"
        Me.ap_codart.OptionsColumn.AllowEdit = False
        Me.ap_codart.OptionsColumn.AllowSort = DevExpress.Utils.DefaultBoolean.[True]
        Me.ap_codart.OptionsColumn.ReadOnly = True
        Me.ap_codart.OptionsFilter.FilterPopupMode = DevExpress.XtraGrid.Columns.FilterPopupMode.CheckedList
        Me.ap_codart.Visible = True
        Me.ap_codart.VisibleIndex = 0
        '
        'xxapx_codart
        '
        Me.xxapx_codart.AppearanceCell.Options.UseBackColor = True
        Me.xxapx_codart.AppearanceCell.Options.UseTextOptions = True
        Me.xxapx_codart.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
        Me.xxapx_codart.Caption = "Descrizione Articolo"
        Me.xxapx_codart.Enabled = False
        Me.xxapx_codart.FieldName = "xxapx_codart"
        Me.xxapx_codart.FilterMode = DevExpress.XtraGrid.ColumnFilterMode.DisplayText
        Me.xxapx_codart.Name = "xxapx_codart"
        Me.xxapx_codart.OptionsColumn.AllowEdit = False
        Me.xxapx_codart.OptionsColumn.AllowSort = DevExpress.Utils.DefaultBoolean.[True]
        Me.xxapx_codart.OptionsColumn.ReadOnly = True
        Me.xxapx_codart.OptionsFilter.FilterPopupMode = DevExpress.XtraGrid.Columns.FilterPopupMode.CheckedList
        Me.xxapx_codart.Visible = True
        Me.xxapx_codart.VisibleIndex = 1
        '
        'xxar_unmis
        '
        Me.xxar_unmis.AppearanceCell.Options.UseBackColor = True
        Me.xxar_unmis.AppearanceCell.Options.UseTextOptions = True
        Me.xxar_unmis.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
        Me.xxar_unmis.Caption = "UMP"
        Me.xxar_unmis.Enabled = False
        Me.xxar_unmis.FieldName = "xxar_unmis"
        Me.xxar_unmis.FilterMode = DevExpress.XtraGrid.ColumnFilterMode.DisplayText
        Me.xxar_unmis.Name = "xxar_unmis"
        Me.xxar_unmis.OptionsColumn.AllowEdit = False
        Me.xxar_unmis.OptionsColumn.AllowSort = DevExpress.Utils.DefaultBoolean.[True]
        Me.xxar_unmis.OptionsColumn.ReadOnly = True
        Me.xxar_unmis.OptionsFilter.FilterPopupMode = DevExpress.XtraGrid.Columns.FilterPopupMode.CheckedList
        Me.xxar_unmis.Visible = True
        Me.xxar_unmis.VisibleIndex = 2
        '
        'ap_esist
        '
        Me.ap_esist.AppearanceCell.Options.UseBackColor = True
        Me.ap_esist.AppearanceCell.Options.UseTextOptions = True
        Me.ap_esist.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
        Me.ap_esist.Caption = "Esistenza"
        Me.ap_esist.Enabled = False
        Me.ap_esist.FieldName = "ap_esist"
        Me.ap_esist.FilterMode = DevExpress.XtraGrid.ColumnFilterMode.DisplayText
        Me.ap_esist.Name = "ap_esist"
        Me.ap_esist.OptionsColumn.AllowEdit = False
        Me.ap_esist.OptionsColumn.AllowSort = DevExpress.Utils.DefaultBoolean.[True]
        Me.ap_esist.OptionsColumn.ReadOnly = True
        Me.ap_esist.OptionsFilter.FilterPopupMode = DevExpress.XtraGrid.Columns.FilterPopupMode.CheckedList
        Me.ap_esist.Visible = True
        Me.ap_esist.VisibleIndex = 3
        '
        'xx_qtadainv
        '
        Me.xx_qtadainv.AppearanceCell.Options.UseBackColor = True
        Me.xx_qtadainv.AppearanceCell.Options.UseTextOptions = True
        Me.xx_qtadainv.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
        Me.xx_qtadainv.Caption = "Qta da inviare"
        Me.xx_qtadainv.Enabled = True
        Me.xx_qtadainv.FieldName = "xx_qtadainv"
        Me.xx_qtadainv.FilterMode = DevExpress.XtraGrid.ColumnFilterMode.DisplayText
        Me.xx_qtadainv.Name = "xx_qtadainv"
        Me.xx_qtadainv.Visible = True
        Me.xx_qtadainv.VisibleIndex = 4
        '
        'xx_dtinv
        '
        Me.xx_dtinv.AppearanceCell.Options.UseBackColor = True
        Me.xx_dtinv.AppearanceCell.Options.UseTextOptions = True
        Me.xx_dtinv.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap
        Me.xx_dtinv.Caption = "Data Invio"
        Me.xx_dtinv.Enabled = True
        Me.xx_dtinv.FieldName = "xx_dtinv"
        Me.xx_dtinv.FilterMode = DevExpress.XtraGrid.ColumnFilterMode.DisplayText
        Me.xx_dtinv.Name = "xx_dtinv"
        Me.xx_dtinv.Visible = True
        Me.xx_dtinv.VisibleIndex = 5
        '
        'grpArtpro
        '
        Me.grpArtpro.AllowDrop = True
        Me.grpArtpro.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpArtpro.Appearance.BackColor = System.Drawing.Color.Transparent
        Me.grpArtpro.Appearance.Options.UseBackColor = True
        Me.grpArtpro.Controls.Add(Me.edEi_codlsar)
        Me.grpArtpro.Controls.Add(Me.lbXxEi_codlsar)
        Me.grpArtpro.Controls.Add(Me.lbEi_codlsar)
        Me.grpArtpro.Controls.Add(Me.lbEi_whereart)
        Me.grpArtpro.Controls.Add(Me.cmdScegliArticoli)
        Me.grpArtpro.Controls.Add(Me.grArtpro)
        Me.grpArtpro.Location = New System.Drawing.Point(4, 42)
        Me.grpArtpro.Name = "grpArtpro"
        Me.grpArtpro.Size = New System.Drawing.Size(612, 456)
        Me.grpArtpro.Text = "GIACENZE DA ESPORTARE"
        '
        'edEi_codlsar
        '
        Me.edEi_codlsar.Location = New System.Drawing.Point(104, 48)
        Me.edEi_codlsar.Name = "edEi_codlsar"
        Me.edEi_codlsar.Properties.Appearance.Options.UseTextOptions = True
        Me.edEi_codlsar.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        Me.edEi_codlsar.Properties.AppearanceDisabled.ForeColor = System.Drawing.Color.Black
        Me.edEi_codlsar.Properties.AppearanceDisabled.Options.UseForeColor = True
        Me.edEi_codlsar.Properties.AutoHeight = False
        Me.edEi_codlsar.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.edEi_codlsar.Properties.MaxLength = 65536
        Me.edEi_codlsar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.edEi_codlsar.Size = New System.Drawing.Size(52, 20)
        '
        'lbXxEi_codlsar
        '
        Me.lbXxEi_codlsar.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbXxEi_codlsar.BackColor = System.Drawing.Color.Transparent
        Me.lbXxEi_codlsar.Location = New System.Drawing.Point(160, 48)
        Me.lbXxEi_codlsar.Name = "lbXxEi_codlsar"
        Me.lbXxEi_codlsar.NTSBordeStyle = NTSInformatica.NTSLabel.NTSBorderStyle.FieldDescription
        Me.lbXxEi_codlsar.Size = New System.Drawing.Size(304, 20)
        Me.lbXxEi_codlsar.UseMnemonic = False
        '
        'lbEi_codlsar
        '
        Me.lbEi_codlsar.BackColor = System.Drawing.Color.Transparent
        Me.lbEi_codlsar.Location = New System.Drawing.Point(4, 48)
        Me.lbEi_codlsar.Name = "lbEi_codlsar"
        Me.lbEi_codlsar.NTSBordeStyle = NTSInformatica.NTSLabel.NTSBorderStyle.FieldCaption
        Me.lbEi_codlsar.Size = New System.Drawing.Size(92, 20)
        Me.lbEi_codlsar.Text = "Lista sel. articoli"
        Me.lbEi_codlsar.UseMnemonic = False
        '
        'lbEi_whereart
        '
        Me.lbEi_whereart.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbEi_whereart.BackColor = System.Drawing.Color.Transparent
        Me.lbEi_whereart.Location = New System.Drawing.Point(104, 24)
        Me.lbEi_whereart.Name = "lbEi_whereart"
        Me.lbEi_whereart.NTSBordeStyle = NTSInformatica.NTSLabel.NTSBorderStyle.FieldDescription
        Me.lbEi_whereart.Size = New System.Drawing.Size(360, 20)
        Me.lbEi_whereart.UseMnemonic = False
        '
        'cmdScegliArticoli
        '
        Me.cmdScegliArticoli.Location = New System.Drawing.Point(5, 23)
        Me.cmdScegliArticoli.Name = "cmdScegliArticoli"
        Me.cmdScegliArticoli.Size = New System.Drawing.Size(96, 20)
        Me.cmdScegliArticoli.Text = "Scegli Articoli"
        '
        'FRMZCEXIU
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.ClientSize = New System.Drawing.Size(640, 511)
        Me.Controls.Add(Me.grpArtpro)
        Me.Controls.Add(Me.grpDatiTrasmissione)
        Me.Controls.Add(Me.grpDatiInvio)
        Me.Controls.Add(Me.grpDatiAccount)
        Me.Controls.Add(Me.barDockControlLeft)
        Me.Controls.Add(Me.barDockControlRight)
        Me.Controls.Add(Me.barDockControlBottom)
        Me.Controls.Add(Me.barDockControlTop)
        Me.Name = "FRMZCEXIU"
        Me.Text = "EXPORT FILE"
        CType(Me.dttSmartArt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NtsBarManager1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpDatiAccount, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpDatiAccount.ResumeLayout(False)
        CType(Me.edEi_password.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.edEi_username.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpDatiInvio, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpDatiInvio.ResumeLayout(False)
        CType(Me.ckModaleOutlook.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.edEi_emaildest.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.edEi_transmissiondate.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cbEi_fromto.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.edEi_legalname.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpDatiTrasmissione, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpDatiTrasmissione.ResumeLayout(False)
        CType(Me.edEi_codfornlux.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.edEi_auxhead1.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.edEi_transmissionnumber.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.edEi_issuedate.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.edEi_partnerid.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grArtpro, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grvArtpro, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpArtpro, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpArtpro.ResumeLayout(False)
        CType(Me.edEi_codlsar.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Public Overridable Sub InitControls()
        Dim dttDirezioni As New DataTable
        Dim dttAuxhead1 As New DataTable
        InitControlsBeginEndInit(Me, False)
        Try
            '-------------------------------------------------
            'carico le immagini della toolbar
            Try
                tlbStampa.Glyph = Bitmap.FromFile(oApp.ChildImageDir & "\print.png")
                tlbStampaVideo.Glyph = Bitmap.FromFile(oApp.ChildImageDir & "\prnscreen.png")
                tlbGuida.Glyph = Bitmap.FromFile(oApp.ChildImageDir & "\help.png")
                tlbEsci.Glyph = Bitmap.FromFile(oApp.ChildImageDir & "\exit.png")
                tlbElabora.Glyph = Bitmap.FromFile(oApp.ChildImageDir & "\elabora.png")
                tlbSalva.Glyph = Bitmap.FromFile(oApp.ChildImageDir & "\save.png")
            Catch ex As Exception
                'non gestisco l'errore: se non c'è una immagine prendo quella standard
            End Try


            grvArtpro.NTSSetParam(oMenu, oApp.Tr(Me, 132186424492455852, "Griglia Articoli"))
            edEi_codlsar.NTSSetParamTabe(oMenu, oApp.Tr(Me, 132186424515281911, "Lista Selezionata Articoli"), tablsar)
            edEi_auxhead1.NTSSetParam(oApp.Tr(Me, 132186424515304467, "Riservati A"))
            edEi_transmissionnumber.NTSSetParam(oMenu, oApp.Tr(Me, 132186424515310113, "Progressivo"), "0")
            edEi_issuedate.NTSSetParam(oMenu, oApp.Tr(Me, 132186424515321300, "Data Creazione"), False, "X", New Date(1900, 1, 1), New Date(2099, 12, 31))
            edEi_partnerid.NTSSetParam(oMenu, oApp.Tr(Me, 132186424515321442, "ID Partner"), 0)
            edEi_emaildest.NTSSetParam(oMenu, oApp.Tr(Me, 132186424515327000, "Destinatario"), 0)
            edEi_transmissiondate.NTSSetParam(oMenu, oApp.Tr(Me, 132186424515332645, "Data Trasmissione"), False, "X", New Date(1900, 1, 1), New Date(2099, 12, 31))
            cbEi_fromto.NTSSetParam(oApp.Tr(Me, 132186424515332789, "From/To"))
            edEi_legalname.NTSSetParam(oMenu, oApp.Tr(Me, 132186424515338463, "Azienda"), 0)
            edEi_password.NTSSetParam(oMenu, oApp.Tr(Me, 132186424515344222, "Password"), 0)
            edEi_username.NTSSetParam(oMenu, oApp.Tr(Me, 132186424515349409, "Username"), 0)
            ap_codart.NTSSetParamSTRTabe(oMenu, oApp.Tr(Me, 132186487538088706, "Articolo"), tabartico, False)
            xxapx_codart.NTSSetParamSTR(oMenu, oApp.Tr(Me, 132186487538094334, "Descrizione Articolo"), 0, True)
            xxar_unmis.NTSSetParamSTR(oMenu, oApp.Tr(Me, 132186487538094411, "UMP"), 0, True)
            ap_esist.NTSSetParamNUM(oMenu, oApp.Tr(Me, 132186487538100146, "Esistenza"), oApp.FormatQta, 15)
            xx_qtadainv.NTSSetParamNUM(oMenu, oApp.Tr(Me, 129663643960156250, "Quantità Da Inviare"), oApp.FormatQta, 15)
            xx_dtinv.NTSSetParamDATA(oMenu, oApp.Tr(Me, 138787412078455196, "Data Invio"), False)



            edEi_codfornlux.NTSSetParam(oMenu, oApp.Tr(Me, 132186509946684332, "cod. forn. lux"), 0)
      ckModaleOutlook.NTSSetParam(oMenu, oApp.Tr(Me, 132246750037760041, "Visualizza prima di inviare"), "S", "N")

      dttDirezioni.Columns.Add("cod", GetType(String)).DefaultValue = ""
            dttDirezioni.Columns.Add("val", GetType(String)).DefaultValue = ""
            dttDirezioni.Rows.Add("toM2", "In uscita")

            cbEi_fromto.DataSource = dttDirezioni
            cbEi_fromto.ValueMember = "cod"
            cbEi_fromto.DisplayMember = "val"


            dttAuxhead1.Columns.Add("cod", GetType(String)).DefaultValue = ""
            dttAuxhead1.Columns.Add("val", GetType(String)).DefaultValue = ""
            dttAuxhead1.Rows.Add("0009", "Luxottica srl")
            dttAuxhead1.Rows.Add("0010", "Luxottica group")
            dttAuxhead1.Rows.Add("0053", "Tristar")
            dttAuxhead1.Rows.Add("0065", "LHKS")

            edEi_auxhead1.DataSource = dttAuxhead1
            edEi_auxhead1.ValueMember = "cod"
            edEi_auxhead1.DisplayMember = "val"


            '-------------------------------------------------
            'chiamo lo script per inizializzare i controlli caricati con source ext
            NTSScriptExec("InitControls", Me, Nothing)
        Catch ex As Exception
            '-------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '-------------------------------------------------
        End Try
        InitControlsBeginEndInit(Me, True)
    End Sub

    Public Overridable Sub Bindcontrols()
        Try
            '-------------------------------------------------
            'se i controlli erano già stati precedentemente collegati, li scollego
            NTSFormClearDataBinding(Me)

            '-------------------------------------------------
            'collego il BindingSource ai vari controlli 

            edEi_auxhead1.NTSDbField = "IMPOSTAZIONI.ei_auxhead1"
            edEi_transmissionnumber.NTSDbField = "IMPOSTAZIONI.ei_transmissionnumber"
            edEi_issuedate.NTSDbField = "IMPOSTAZIONI.ei_issuedate"
            edEi_partnerid.NTSDbField = "IMPOSTAZIONI.ei_partnerid"
            edEi_emaildest.NTSDbField = "IMPOSTAZIONI.ei_emaildest"
            edEi_transmissiondate.NTSDbField = "IMPOSTAZIONI.ei_transmissiondate"
            cbEi_fromto.NTSDbField = "IMPOSTAZIONI.ei_fromto"
            edEi_legalname.NTSDbField = "IMPOSTAZIONI.ei_legalname"
            edEi_password.NTSDbField = "IMPOSTAZIONI.ei_password"
            edEi_username.NTSDbField = "IMPOSTAZIONI.ei_username"
            edEi_codlsar.NTSDbField = "IMPOSTAZIONI.ei_codlsar"
            lbXxEi_codlsar.NTSDbField = "IMPOSTAZIONI.xxei_codlsar"
            lbEi_whereart.NTSDbField = "IMPOSTAZIONI.ei_whereart"
            edEi_codfornlux.NTSDbField = "IMPOSTAZIONI.ei_codfornlux"

            dcEXIU.DataSource = dsEXIU.Tables("IMPOSTAZIONI")
            dcArtpro.DataSource = dsEXIU.Tables("ARTPROX")

            grArtpro.DataSource = dcArtpro

            dsEXIU.AcceptChanges()

            '-------------------------------------------------
            'per agganciare al dataset i vari controlli
            NTSFormAddDataBinding(dcEXIU, Me)

        Catch ex As Exception
            '-------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '-------------------------------------------------
        End Try
    End Sub

#Region "Toolbar"
    Public Overridable Sub tlbGuida_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles tlbGuida.ItemClick
        SendKeys.SendWait("{F1}")
    End Sub

    Public Overridable Sub tlbEsci_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles tlbEsci.ItemClick
        Try
            Me.Close()
        Catch ex As Exception
            '-------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '-------------------------------------------------
        End Try
    End Sub

    Public Overridable Sub cmdScegliArticoli_Click(sender As Object, e As EventArgs) Handles cmdScegliArticoli.Click
        Dim oPar As CLE__CLDP = Nothing
        Try
            oPar = New CLE__CLDP
            oPar.bVisGriglia = False
            oPar.strTipoArticolo = "N"
            NTSZOOM.strIn = ""
            NTSZOOM.ZoomStrIn("ZOOMARTICO", DittaCorrente, oPar)
            If oPar.CANCELZOOM = False And oPar.strOut <> "" Then dsEXIU.Tables("IMPOSTAZIONI").Rows(0)!ei_whereart = oPar.strOut.Trim

        Catch ex As Exception
            '-------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '-------------------------------------------------	
        End Try
    End Sub

    Public Overridable Sub tlbSalva_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles tlbSalva.ItemClick
        Try
            If oCleEXIU.SalvaImpostazioni(DittaCorrente, dsEXIU) Then
                oApp.MsgBoxInfo("Impostazioni salvate.")
            End If
        Catch ex As Exception
            '-------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '-------------------------------------------------	
        End Try
    End Sub

    Private Sub tlbElabora_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles tlbElabora.ItemClick
        Try
            If oCleEXIU.Elabora(ckModaleOutlook.Checked) Then
                oApp.MsgBoxInfo("File generato correttamente.")
            End If
        Catch ex As Exception
            '-------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '-------------------------------------------------	
        End Try
    End Sub

    Private Sub FRMZCEXIU_ActivatedFirst(sender As Object, e As EventArgs) Handles MyBase.ActivatedFirst
        Try
            If oApp.Batch Then
                Me.Top = 0
                Me.Left = 0
                ckModaleOutlook.Checked = False
                tlbElabora_ItemClick(Nothing, Nothing)
                Me.Close()
            End If
        Catch ex As Exception
            '-------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '-------------------------------------------------	
        End Try
    End Sub

#End Region

End Class
