Public Class ssSetting
    Public Sub Value(ByRef parmSetting As Object, ByVal parmValue As Object)

        'Boolean
        If TypeOf parmSetting Is Boolean Then
            If TypeOf parmValue Is Boolean Then
                parmSetting = parmValue
            End If
        End If

        'Number
        If TypeOf parmSetting Is Integer Then
            If TypeOf parmValue Is Integer Then
                parmSetting = parmValue
            End If
        End If

        'String
        If TypeOf parmSetting Is String Then
            If TypeOf parmValue Is String Then
                parmSetting = parmValue
            End If
        End If

        'parmSetting = parmValue
        My.Settings.Save()

    End Sub

    Public Sub SetValue(ByRef parmSetting As Object, ByVal parmValue As Object)

        'Boolean
        If TypeOf parmSetting Is Boolean Then
            If TypeOf parmValue Is Boolean Then
                parmSetting = parmValue
            End If
        End If

        'Number
        If TypeOf parmSetting Is Integer Then
            If TypeOf parmValue Is Integer Then
                parmSetting = parmValue
            End If
        End If

        'String
        If TypeOf parmSetting Is String Then
            If TypeOf parmValue Is String Then
                parmSetting = parmValue
            End If
        End If

        'parmSetting = parmValue
        My.Settings.Save()

    End Sub


End Class
