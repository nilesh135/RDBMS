Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.ComponentModel
Imports Microsoft.Data.ConnectionUI
Imports System.Data.SqlClient

Friend Class SQLConnectionDialog
    Private cp As SqlFileConnectionProperties
    Private uic As SqlConnectionUIControl

    'Allows the user to change the title of the dialog
    Public Overrides Property Text() As String
        Get
            Return MyBase.Text
        End Get
        Set(ByVal value As String)
            MyBase.Text = value
        End Set
    End Property

    'Pass the original connection string or get the resulting connection string
    Public Property ConnectionString() As String
        Get
            Return cp.ConnectionStringBuilder.ConnectionString
        End Get
        Set(ByVal value As String)
            cp.ConnectionStringBuilder.ConnectionString = value
        End Set
    End Property


    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub Dialog1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Padding = New Padding(5)
        Dim adv As Button = New Button
        Dim Tst As Button = New Button

        'Size the form and place the uic, Test connection button, and advanced button
        uic.LoadProperties()
        uic.Dock = DockStyle.Top
        uic.Parent = Me
        Me.ClientSize = Size.Add(uic.MinimumSize, New Size(10, (adv.Height + 25)))
        Me.MinimumSize = Me.Size
        With adv
            .Text = "Advanced"
            .Dock = DockStyle.None
            .Location = New Point((uic.Width - .Width), (uic.Bottom + 10))
            .Anchor = (AnchorStyles.Right Or AnchorStyles.Top)
            AddHandler .Click, AddressOf Me.Advanced_Click
            .Parent = Me
        End With

        With Tst
            .Text = "Test Connection"
            .Width = 100
            .Dock = DockStyle.None
            .Location = New Point((uic.Width - .Width) - adv.Width - 10, (uic.Bottom + 10))
            .Anchor = (AnchorStyles.Right Or AnchorStyles.Top)
            AddHandler .Click, AddressOf Me.Test_Click
            .Parent = Me
        End With
    End Sub

    Private Sub Advanced_Click(ByVal sender As Object, ByVal e As EventArgs)
        'Set up a form to display the advanced connection properties
        Dim frm As Form = New Form
        Dim pg As PropertyGrid = New PropertyGrid
        pg.SelectedObject = cp
        pg.Dock = DockStyle.Fill
        pg.Parent = frm
        frm.ShowDialog()
    End Sub

    Private Sub Test_Click(ByVal sender As Object, ByVal e As EventArgs)
        'Test the connection
        Dim conn As New SqlConnection()
        conn.ConnectionString = cp.ConnectionStringBuilder.ConnectionString
        Try
            conn.Open()
            MsgBox("Test Connection Succeeded.", MsgBoxStyle.Exclamation)
        Catch ex As Exception
            MsgBox("Test Connection Failed.", MsgBoxStyle.Critical)
        Finally
            Try
                conn.Close()
            Catch ex As Exception
                
            End Try
        End Try

    End Sub
End Class
