Imports System.Configuration
Public Class SQL_Connection_Dialog

    Private _Frm_SQLConnectionDialog As SQLConnectionDialog

    Public Sub New()
        MyBase.New()
        _Frm_SQLConnectionDialog = New SQLConnectionDialog
    End Sub

    Public Property Title() As String
        Get
            Return Me._Frm_SQLConnectionDialog.Text
        End Get
        Set(ByVal value As String)
            Me._Frm_SQLConnectionDialog.Text = value
        End Set
    End Property

    Public Property ConnectionString() As String
        Get
            Return Me._Frm_SQLConnectionDialog.ConnectionString
        End Get
        Set(ByVal value As String)
            Me._Frm_SQLConnectionDialog.ConnectionString = value
        End Set
    End Property


    Public Sub SaveChange_To_App_Config(ByVal connectionName As String)
        Dim Config As Configuration
        Dim Section As ConnectionStringsSection
        Dim Setting As ConnectionStringSettings
        Dim ConnectionFullName As String

        'There is no inbuilt way to change application setting values in the config file.
        'So that needs to be done manually by calling config section object.

        Try
            'Concatenate the full settings name
            'This differs from Jakob Lithner. Runtime Connection Wizard
            'The ConnectionFullName needs to refer to the Assembly calling this DLL
            ConnectionFullName = String.Format("{0}.MySettings.{1}", System.Reflection.Assembly.GetCallingAssembly.EntryPoint.DeclaringType.Namespace, connectionName)

            'Point out the objects to manipulate
            Config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
            Section = CType(Config.GetSection("connectionStrings"), ConnectionStringsSection)
            Setting = Section.ConnectionStrings(ConnectionFullName)

            'Ensure connection setting is defined (Note: A default value must be set to save the connection setting!)
            If IsNothing(Setting) Then Throw New Exception("There is no connection with this name defined in the config file.")

            'Set value and save it to the config file
            'This differs from Jakob Lithner. Runtime Connection Wizard
            'We only want to save the modified portion of the config file 
            Setting.ConnectionString = Me.ConnectionString
            Config.Save(ConfigurationSaveMode.Modified, True)

            
        Catch ex As Exception

        End Try
    End Sub

    Public Function ShowDialog() As System.Windows.Forms.DialogResult
        Return Me._Frm_SQLConnectionDialog.ShowDialog

    End Function

End Class
