Imports System
Imports System.Runtime.InteropServices
Imports Microsoft.Win32

Public Class frmInstall
    Inherits System.Windows.Forms.Form
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents cmdInstall As System.Windows.Forms.Button
    Friend WithEvents FileWatcher As System.IO.FileSystemWatcher
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmInstall))
        Me.cmdInstall = New System.Windows.Forms.Button()
        Me.FileWatcher = New System.IO.FileSystemWatcher()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        CType(Me.FileWatcher, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdInstall
        '
        Me.cmdInstall.BackColor = System.Drawing.Color.Silver
        Me.cmdInstall.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.cmdInstall.FlatAppearance.MouseDownBackColor = System.Drawing.Color.RoyalBlue
        Me.cmdInstall.FlatAppearance.MouseOverBackColor = System.Drawing.Color.RoyalBlue
        Me.cmdInstall.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdInstall.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdInstall.ForeColor = System.Drawing.Color.White
        Me.cmdInstall.Location = New System.Drawing.Point(14, 201)
        Me.cmdInstall.Name = "cmdInstall"
        Me.cmdInstall.Size = New System.Drawing.Size(195, 54)
        Me.cmdInstall.TabIndex = 1
        Me.cmdInstall.Text = "Instalar"
        Me.cmdInstall.UseVisualStyleBackColor = False
        '
        'FileWatcher
        '
        Me.FileWatcher.EnableRaisingEvents = True
        Me.FileWatcher.SynchronizingObject = Me
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.cmdCancel.FlatAppearance.MouseDownBackColor = System.Drawing.Color.RoyalBlue
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.RoyalBlue
        Me.cmdCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.White
        Me.cmdCancel.Location = New System.Drawing.Point(272, 201)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(195, 54)
        Me.cmdCancel.TabIndex = 8
        Me.cmdCancel.Text = "Cancelar"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox1.ForeColor = System.Drawing.Color.Black
        Me.TextBox1.Location = New System.Drawing.Point(7, 22)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(438, 22)
        Me.TextBox1.TabIndex = 10
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TextBox1)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox1.ForeColor = System.Drawing.Color.SteelBlue
        Me.GroupBox1.Location = New System.Drawing.Point(14, 104)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(453, 58)
        Me.GroupBox1.TabIndex = 13
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Instalación en el directorio:"
        '
        'frmInstall
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(497, 327)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdInstall)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmInstall"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Instalador Add-On"
        Me.TopMost = True
        CType(Me.FileWatcher, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Data members"
    'Curso SDK
    Private sAddonName As String = "BackOrder"
    Private sAddonEXE As String = "Permisos" '.exe
    Private sAddonConfig As String = "BackOrder"
    Private strDll As String ' The path of "AddOnInstallAPI.dll"
    Private strDest As String ' Installation target path
    Private bFileCreated As Boolean ' True if the file was created
#End Region

#Region "Declarations"
    ' Declaring the functions inside "AddOnInstallAPI.dll"

    'EndInstall - Signals SBO that the installation is complete.
    Declare Function EndInstallEx Lib "AddOnInstallAPI.dll" (ByVal str As String, ByVal b As Boolean) As Int32
    'EndUnInstall - Signals SBO that the uninstallation is complete.
    Declare Function EndUninstall Lib "AddOnInstallAPI.dll" (ByVal str As String, ByVal b As Boolean) As Int32
    'SetAddOnFolder - Use it if you want to change the installation folder.
    Declare Function SetAddOnFolder Lib "AddOnInstallAPI.dll" (ByVal srrPath As String) As Int32
    'RestartNeeded - Use it if your installation requires a restart, it will cause
    'the SBO application to close itself after the installation is complete.
    Declare Function RestartNeeded Lib "AddOnInstallAPI.dll" () As Int32
    'the SBO application to close itself after the installation is complete.
    Declare Function B1Info Lib "AddOnInstallAPI.dll" (ByRef lpBuffer As String, ByVal length As Int32) As Int32

#End Region

#Region "Methods"

    ' Read the addon path from the registry
    Public Function ReadPath() As String
        Dim sAns As String
        Dim sErr As String = "BackOrder"

        sAns = RegValue(RegistryHive.LocalMachine, "SOFTWARE", sAddonName, sErr)
        ReadPath = sAns
        If Not (sAns <> "") Then
            MessageBox.Show("Error al registrar el Add-On: " & sErr)
        End If
    End Function

    ' This Function reads values to the registry
    Public Function RegValue(ByVal Hive As RegistryHive,
          ByVal Key As String, ByVal ValueName As String,
          Optional ByRef ErrInfo As String = "") As String

        Dim objParent As RegistryKey
        Dim objSubkey As RegistryKey
        Dim sAns As String = ""
        Select Case Hive
            Case RegistryHive.ClassesRoot
                objParent = Registry.ClassesRoot
            Case RegistryHive.CurrentConfig
                objParent = Registry.CurrentConfig
            Case RegistryHive.CurrentUser
                objParent = Registry.CurrentUser
            Case RegistryHive.DynData
                objParent = Registry.DynData
            Case RegistryHive.LocalMachine
                objParent = Registry.LocalMachine
            Case RegistryHive.PerformanceData
                objParent = Registry.PerformanceData
            Case RegistryHive.Users
                objParent = Registry.Users

        End Select

        Try
            objSubkey = objParent.OpenSubKey(Key)
            'if can't be found, object is not initialized
            If Not objSubkey Is Nothing Then
                sAns = (objSubkey.GetValue(ValueName))
            End If

        Catch ex As Exception

            ErrInfo = ex.Message
        Finally

            'if no error but value is empty, populate errinfo
            If ErrInfo = "" And sAns = "" Then
                ErrInfo =
                   "No value found for requested registry key"
            End If
        End Try
        Return sAns
    End Function

    ' This Function writes values to the registry
    Public Function WriteToRegistry(ByVal _
    ParentKeyHive As RegistryHive,
    ByVal SubKeyName As String,
    ByVal ValueName As String,
    ByVal Value As Object) As Boolean

        Dim objSubKey As RegistryKey
        Dim sException As String = ""
        Dim objParentKey As RegistryKey
        Dim bAns As Boolean

        Try

            Select Case ParentKeyHive
                Case RegistryHive.ClassesRoot
                    objParentKey = Registry.ClassesRoot
                Case RegistryHive.CurrentConfig
                    objParentKey = Registry.CurrentConfig
                Case RegistryHive.CurrentUser
                    objParentKey = Registry.CurrentUser
                Case RegistryHive.DynData
                    objParentKey = Registry.DynData
                Case RegistryHive.LocalMachine
                    objParentKey = Registry.LocalMachine
                Case RegistryHive.PerformanceData
                    objParentKey = Registry.PerformanceData
                Case RegistryHive.Users
                    objParentKey = Registry.Users
            End Select

            'Open 
            objSubKey = objParentKey.OpenSubKey(SubKeyName, True)
            'create if doesn't exist
            If objSubKey Is Nothing Then
                objSubKey = objParentKey.CreateSubKey(SubKeyName)
            End If


            objSubKey.SetValue(ValueName, Value)
            bAns = True
        Catch ex As Exception
            bAns = False

        End Try

        Return True

    End Function

    Public Sub DeleteToRegistry(ByVal _
       ParentKeyHive As RegistryHive,
       ByVal SubKeyName As String,
       ByVal ValueName As String)

        Dim objSubKey As RegistryKey
        Dim sException As String = ""
        Dim objParentKey As RegistryKey
        Try
            Select Case ParentKeyHive
                Case RegistryHive.ClassesRoot
                    objParentKey = Registry.ClassesRoot
                Case RegistryHive.CurrentConfig
                    objParentKey = Registry.CurrentConfig
                Case RegistryHive.CurrentUser
                    objParentKey = Registry.CurrentUser
                Case RegistryHive.DynData
                    objParentKey = Registry.DynData
                Case RegistryHive.LocalMachine
                    objParentKey = Registry.LocalMachine
                Case RegistryHive.PerformanceData
                    objParentKey = Registry.PerformanceData
                Case RegistryHive.Users
                    objParentKey = Registry.Users
            End Select
            'Open 
            objSubKey = objParentKey.OpenSubKey(SubKeyName, True)
            'create if doesn't exist
            If objSubKey Is Nothing Then
                objSubKey = objParentKey.CreateSubKey(SubKeyName)
            End If
            objSubKey.DeleteValue(ValueName)
        Catch ex As Exception
        End Try
    End Sub

    ' This function extracts the given add-on into the path specified
    Private Sub ExtractFile(ByVal Path As String, ByVal FileName As String)
        Try
            Dim AddonExeFile As IO.FileStream
            Dim thisExe As System.Reflection.Assembly
            thisExe = System.Reflection.Assembly.GetExecutingAssembly()
            Dim sTargetPath As String = Path & FileName 'path & "\" & sAddonName & "exe"
            Dim sSourcePath As String = Path & FileName & ".tmp"

            If Not IO.Directory.Exists(Path) Then
                IO.Directory.CreateDirectory(Path)
            End If

            Dim file As System.IO.Stream

            file = thisExe.GetManifestResourceStream("Installer." & FileName)

            ' Create a tmp file first, after file is extracted change to exe
            If IO.File.Exists(sSourcePath) Then
                IO.File.Delete(sSourcePath)
            End If
            AddonExeFile = IO.File.Create(sSourcePath)

            Dim buffer() As Byte
            ReDim buffer(file.Length)

            file.Read(buffer, 0, file.Length)
            AddonExeFile.Write(buffer, 0, file.Length)
            AddonExeFile.Close()

            If IO.File.Exists(sTargetPath) Then
                IO.File.Delete(sTargetPath)
            End If
            ' Change file extension to exe
            IO.File.Move(sSourcePath, sTargetPath)
            file.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ' This procedure delets the addon files

    Public Sub DeleteFile(Filename As String)
        ' Delete the addon file
        If IO.File.Exists(Filename) Then
            IO.File.Delete(Filename)
        End If
    End Sub

    Private Sub KillApp(ByVal ProcessName As String)
        For Each prog As Process In Process.GetProcesses
            If prog.ProcessName = ProcessName Then
                prog.Kill()
            End If
        Next
    End Sub

    Private Sub UnInstall()
        Dim path As String
        path = ReadPath() ' Reads the addon path from the registry
        If path <> "" Then
            Try
                ' Delete the addon EXE file
                If IO.Directory.Exists(path & "\") Then
                    KillApp(sAddonEXE)

                    'Curso SDK
                    'DeleteFile(path & "\" & "Permisos.b1s")
                    DeleteFile(path & "\" & sAddonEXE & ".exe")
                    DeleteFile(path & "\" & sAddonEXE & ".exe.config")
                    DeleteFile(path & "\" & "Permisos.pdb")

                    DeleteFile(path & "\" & "Sap.Data.Hana.v4.5.dll")
                    DeleteFile(path & "\" & "SAPBusinessOneSDK.dll")




                    IO.Directory.Delete(path & "\", True)
                    DeleteToRegistry(RegistryHive.LocalMachine, "SOFTWARE", sAddonName)
                    MessageBox.Show("Add-On " & sAddonName & " ha sido desinstalado correctamente!", "3Core", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    EndUninstall("", True) ' Finished Uninstalling
                Else
                    MessageBox.Show(sAddonName & " no se encuentra instalado!", "ERROR AL ELIMINAR ADDON:")
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, "ERROR AL DESINSTALAR:", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
            End Try
        Else
            MessageBox.Show("Ruta de acceso no valida!", "REGISTRO ADDON:")
        End If
        ' Terminate the application
        GC.Collect()
        End
    End Sub

    ' This procedure copies the addon file to the installation folder        
    Private Sub Install()
        Try
            Environment.CurrentDirectory = strDll ' For Dll function calls will work

            If strDest <> TextBox1.Text Then ' Change the installation folder
                SetAddOnFolder(TextBox1.Text)
                strDest = TextBox1.Text
            End If

            If Not (IO.Directory.Exists(strDest)) Then
                IO.Directory.CreateDirectory(strDest) ' Create installation folder
            End If

            FileWatcher.Path = strDest
            FileWatcher.EnableRaisingEvents = True

            'Curso SDK
            'ExtractFile(strDest & "\", "Permisos.b1s")
            ExtractFile(strDest & "\", sAddonEXE & ".exe")
            ExtractFile(strDest & "\", sAddonEXE & ".exe.config")
            ExtractFile(strDest & "\", "Permisos.pdb")

            ExtractFile(strDest & "\", "Sap.Data.Hana.v4.5.dll")
            ExtractFile(strDest & "\", "SAPBusinessOneSDK.dll")



            While bFileCreated = False
                Application.DoEvents()
                'Don't continue running until the file is copied...
            End While

            'If chkRestart.Checked Then
            '    RestartNeeded() ' Inform SBO the restart is needed
            'End If

            EndInstallEx("", True) ' Inform SBO the installation ended successfully
            'Write installation Folder to registry
            Dim bAns As Boolean

            bAns = WriteToRegistry(RegistryHive.LocalMachine, "SOFTWARE", sAddonName, strDest)
            MessageBox.Show("Add-On " & sAddonName & " ha sido instalado correctamente!", "Instalador Add-On", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Windows.Forms.Application.Exit() ' Exit the installer
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "Addon Installer")
        End Try
    End Sub

#End Region

#Region "Events"

    Private Sub frmInstall_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Text = " Instalación del AddOn " & sAddonName & " "
            'Dim strAppPath As String

            ' The command line parameters, seperated by '|' will be broken to this array
            Dim strCmdLineElements(2) As String

            Dim strCmdLine As String ' The whole command line

            Dim NumOfParams As Integer 'The number of parameters in the command line (should be 2)


            NumOfParams = Environment.GetCommandLineArgs.Length

            If NumOfParams = 2 Then
                strCmdLine = Environment.GetCommandLineArgs.GetValue(1)
                If strCmdLine.ToUpper = "/U" Then
                    UnInstall()
                End If
                strCmdLineElements = strCmdLine.Split("|")

                ' Get Install destination Folder
                strDest = strCmdLineElements.GetValue(0)
                TextBox1.Text = strDest
                ' Get the "AddOnInstallAPI.dll" path
                strDll = strCmdLineElements.GetValue(1)
                strDll = strDll.Remove((strDll.Length - 19), 19) ' Only the path is needed
            Else
                MessageBox.Show("Este instalador debe ser ejecutado desde SAP Business One",
                                "Instalador Add-On", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Windows.Forms.Application.Exit()
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub cmdInstall_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInstall.Click
        Install()
    End Sub

    Private Sub chkDefaultFolder_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'TextBox1.Enabled = Not (chkDefaultFolder.Checked)
        'TextBox1.ReadOnly = True
    End Sub

    ' This event happens when the addon exe file is renamed to extention
    Private Sub FileWatcher_Renamed(ByVal sender As Object, ByVal e As System.IO.RenamedEventArgs) Handles FileWatcher.Renamed
        bFileCreated = True
        FileWatcher.EnableRaisingEvents = False
    End Sub

    Public Sub ShowError(ByVal ex As Exception)
        MsgBox(ex.Message & vbNewLine & "Source:" & ex.StackTrace, MsgBoxStyle.Information, "Addon Installer")
    End Sub

#End Region

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        EndInstallEx("", False) ' Inform SBO the installation was not complete
        'MessageBox.Show("Installation was canceled", "Installation was canceled", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Windows.Forms.Application.Exit() ' Exit the installer
    End Sub

End Class

