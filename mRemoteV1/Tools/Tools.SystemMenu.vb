Namespace Tools
    Public Class SystemMenu
        Public Enum Flags
            MF_STRING = App.NativeMethods.MF_STRING
            MF_SEPARATOR = App.NativeMethods.MF_SEPARATOR
            MF_BYCOMMAND = App.NativeMethods.MF_BYCOMMAND
            MF_BYPOSITION = App.NativeMethods.MF_BYPOSITION
            MF_POPUP = App.NativeMethods.MF_POPUP

            WM_SYSCOMMAND = App.NativeMethods.WM_SYSCOMMAND
        End Enum

        Public SystemMenuHandle As IntPtr
        Public FormHandle As IntPtr

        Public Sub New(ByVal Handle As IntPtr)
            FormHandle = Handle
            SystemMenuHandle = App.NativeMethods.GetSystemMenu(FormHandle, False)
        End Sub

        Public Sub Reset()
            SystemMenuHandle = App.NativeMethods.GetSystemMenu(FormHandle, True)
        End Sub

        Public Sub AppendMenuItem(ByVal ParentMenu As IntPtr, ByVal Flags As Flags, ByVal ID As Integer, ByVal Text As String)
            App.NativeMethods.AppendMenu(ParentMenu, Flags, ID, Text)
        End Sub

        Public Function CreatePopupMenuItem() As IntPtr
            Return App.NativeMethods.CreatePopupMenu()
        End Function

        Public Function InsertMenuItem(ByVal SysMenu As IntPtr, ByVal Position As Integer, ByVal Flags As Flags, ByVal Submenu As IntPtr, ByVal Text As String) As Boolean
            Return App.NativeMethods.InsertMenu(SysMenu, Position, Flags, Submenu, Text)
        End Function

        Public Function SetBitmap(ByVal Menu As IntPtr, ByVal Position As Integer, ByVal Flags As Flags, ByVal Bitmap As Bitmap) As IntPtr
            Return App.NativeMethods.SetMenuItemBitmaps(Menu, Position, Flags, Bitmap.GetHbitmap(), Bitmap.GetHbitmap)
        End Function
    End Class
End Namespace