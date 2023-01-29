Public Class frmProblemasGeodesicos
    'Variaveis para o Problema Directo'
    Dim a, f, e2, lat1Dec, lat1Rad, lat2Rad, long1Dec, long1Rad, az12Dec, az12Rad, N1, N2,
        M1, B, C, D, elemE, h, s12, difLatSeg, varLatSeg, varLongSeg, lat2Dec, T12, long2Dec As Double
    'Variaveis para o Problema Inverso'
    Dim a_PI, f_PI, e2_PI, lat1Dec_PI, lat1Rad_PI, lat2Rad_PI, long1Dec_PI, long1Rad_PI, az12Dec_PI, az12Rad_PI,
        elemF_PI, difLatSeg_PI, varLatPI, varLongPI, lat2Dec_PI, long2Dec_PI, long2Rad_PI, N_m, M_m, N1_PI, M1_PI, N2_PI, M2_PI,
        latM, x, y As Double

    'INICIO DO PROBLEMA DICRECTO'
    Private Sub btnResolver_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnResolver.Click
        a = Val(txtSemiEixoMaior.Text)
        f = 1 / Val(txtInvAchatamento.Text)
        e2 = 2 * f - f * f
        txtPrimExc.Text = e2

        'Entrada da Latitude 1 e calculo em decimal e radianos'
        lat1Dec = Val(txtLat1G.Text) + Val(txtLat1M.Text) / 60 + Val(txtLat1S.Text) / 3600
        txtLat1Dec.Text = lat1Dec
        lat1Rad = lat1Dec * Math.PI / 180
        txtLat1Rad.Text = lat1Rad

        'Entrada da Longiude 1 e calculo em decimal e radianos'
        long1Dec = Val(txtLong1G.Text) + Val(txtLong1M.Text) / 60 + Val(txtLong1S.Text) / 3600
        txtLong1Dec.Text = long1Dec
        long1Rad = long1Dec * Math.PI / 180
        txtLong1Rad.Text = long1Rad

        'Entrada do Azimute e calculo em decimal e radianos'
        az12Dec = Val(txtAz12G.Text) + Val(txtAz12M.Text) / 60 + Val(txtAz12S.Text) / 3600
        txtAz12Dec.Text = az12Dec
        az12Rad = az12Dec * Math.PI / 180
        txtAz12Rad.Text = az12Rad

        N1 = a / Math.Sqrt(1 - e2 * Math.Pow(Math.Sin(lat1Rad), 2))
        M1 = (a * (1 - e2)) / Math.Sqrt(Math.Pow(1 - e2 * Math.Pow(Math.Sin(lat1Rad), 2), 3))
        txtM1.Text = M1
        txtN1.Text = N1
        B = 648000 / (M1 * Math.PI)
        C = -1 * 648000 * Math.Tan(lat1Rad) / (2 * M1 * N1 * Math.PI)
        D = -1 * (3 * e2 * Math.Cos(lat1Rad) * Math.Sin(lat1Rad) * Math.PI) / (2 * 648000 * (1 - e2 * Math.Pow(Math.Sin(lat1Rad), 2)))
        elemE = (1 + 3 * Math.Pow(Math.Tan(lat1Rad), 2)) / (6 * Math.Pow(N1, 2))
        s12 = Val(txtDist12.Text)
        h = 648000 * s12 * Math.Cos(az12Rad) / (M1 * Math.PI)
        difLatSeg = B * s12 * Math.Cos(az12Rad) - C * Math.Pow(s12, 2) * Math.Pow(Math.Sin(az12Rad), 2) - h * elemE * Math.Pow(s12, 2) * Math.Pow(Math.Sin(az12Rad), 2)
        varLatSeg = difLatSeg - D * Math.Pow(difLatSeg, 2)

        'Calculo da Latitude 2 em Graus,Minutos, Segundos'
        lat2Dec = -lat1Dec + varLatSeg / 3600

        If lat2Dec < 0 Then
            lat2Dec = -1 * lat2Dec
        End If
        While lat2Dec > 360
            lat2Dec = lat2Dec - 360
        End While
        lat2Rad = lat2Dec * Math.PI / 180

        N2 = a / Math.Sqrt(1 - e2 * Math.Pow(Math.Sin(lat2Rad), 2))
        T12 = s12 * Math.Sin(az12Rad) / (N2 * Math.Cos(lat2Rad))
        varLongSeg = (648000 * T12 / Math.PI) * (1 - (Math.Pow(s12, 2) / (6 * Math.Pow(N2, 2))) + (Math.Pow(T12, 2) / 6))
        long2Dec = -long1Dec + varLongSeg / 3600

        'Calculo da Longitude 2 em Graus,Minutos, Segundos'
        If long2Dec < 0 Then
            long2Dec = -1 * long2Dec
        End If
        While long2Dec > 360
            long2Dec = long2Dec - 360
        End While

        txtElementoB.Text = B
        txtElementoC.Text = C
        txtElementoD.Text = D
        txtElementoE.Text = elemE
        txtElemento_h.Text = h
        txtDifLat.Text = difLatSeg
        txtVarLat.Text = varLatSeg
        txtLat2G.Text = Int(lat2Dec)
        txtLat2M.Text = Int((lat2Dec - Int(lat2Dec)) * 60)
        txtLat2S.Text = Math.Round((((lat2Dec - Int(lat2Dec)) * 60) - Int((lat2Dec - Int(lat2Dec)) * 60)) * 60, 6)
        txtN2.Text = N2
        txtElementoT.Text = T12
        txtVarLong.Text = varLongSeg
        txtLong2G.Text = Int(long2Dec)
        txtLong2M.Text = Int((long2Dec - Int(long2Dec)) * 60)
        txtLong2S.Text = Math.Round((((long2Dec - Int(long2Dec)) * 60) - Int((long2Dec - Int(long2Dec)) * 60)) * 60, 6)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        txtSemiEixoMaior.Text = 6378160
        txtInvAchatamento.Text = 298.25
        txtLat1G.Text = 25
        txtLat1M.Text = 22
        txtLat1S.Text = 49.321
        txtLong1G.Text = 49
        txtLong1M.Text = 27
        txtLong1S.Text = 34.245
        txtAz12G.Text = 245
        txtAz12M.Text = 26
        txtAz12S.Text = 32.453
        txtDist12.Text = 35522.453
    End Sub

    Private Sub btnLimpar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLimpar.Click
        txtSemiEixoMaior.Text = ""
        txtInvAchatamento.Text = ""
        txtLat1G.Text = ""
        txtLat1M.Text = ""
        txtLat1S.Text = ""
        txtLong1G.Text = ""
        txtLong1M.Text = ""
        txtLong1S.Text = ""
        txtAz12G.Text = ""
        txtAz12M.Text = ""
        txtAz12S.Text = ""
        txtDist12.Text = ""
        txtElementoB.Text = ""
        txtElementoC.Text = ""
        txtElementoD.Text = ""
        txtElementoE.Text = ""
        txtElemento_h.Text = ""
        txtDifLat.Text = ""
        txtVarLat.Text = ""
        txtLat2G.Text = ""
        txtLat2M.Text = ""
        txtLat2S.Text = ""
        txtN2.Text = ""
        txtElementoT.Text = ""
        txtVarLong.Text = ""
        txtLong2G.Text = ""
        txtLong2M.Text = ""
        txtLong2S.Text = ""
        txtPrimExc.Text = ""
        txtLat1Dec.Text = ""
        txtLat1Rad.Text = ""
        txtLong1Dec.Text = ""
        txtLong1Rad.Text = ""
        txtAz12Dec.Text = ""
        txtAz12Rad.Text = ""
        txtM1.Text = ""
        txtN1.Text = ""
        'FIM DO PROBLEMA DICRECTO'
    End Sub

    'INICIO DO PROBLEMA INVERSO'
    Private Sub btnResolver2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnResolver2.Click
        a_PI = Val(txtSemiEixoMaiorPI.Text)
        f_PI = 1 / Val(txtInvAchatamentoPI.Text)
        e2_PI = 2 * f_PI - f_PI * f_PI
        txtPrimExcPI.Text = e2_PI

        lat1Dec_PI = Val(txtLat1GPI.Text) + Val(txtLat1MPI.Text) / 60 + Val(txtLat1SPI.Text) / 3600
        txtLat1DecPI.Text = lat1Dec_PI
        lat1Rad_PI = lat1Dec_PI * Math.PI / 180
        txtLat1RadPI.Text = lat1Rad_PI

        long1Dec_PI = Val(txtLong1GPI.Text) + Val(txtLong1MPI.Text) / 60 + Val(txtLong1SPI.Text) / 3600
        txtLong1DecPI.Text = long1Dec_PI
        long1Rad_PI = long1Dec_PI * Math.PI / 180
        txtLong1RadPI.Text = long1Rad_PI


        lat2Dec_PI = Val(txtLat2GPI.Text) + Val(txtLat2MPI.Text) / 60 + Val(txtLat2SPI.Text) / 3600
        txtLat2DecPI.Text = lat2Dec_PI
        lat2Rad_PI = lat2Dec_PI * Math.PI / 180
        txtLat2RadPI.Text = lat2Rad_PI

        long2Dec_PI = Val(txtLong2GPI.Text) + Val(txtLong2MPI.Text) / 60 + Val(txtLong2SPI.Text) / 3600
        txtLong2DecPI.Text = long2Dec_PI
        long2Rad_PI = long2Dec_PI * Math.PI / 180
        txtLong2RadPI.Text = long2Rad_PI

        'Determinacao de M_m, N_m'
        N1_PI = a_PI / Math.Sqrt(1 - e2_PI * Math.Pow(Math.Sin(lat1Rad_PI), 2))
        M1_PI = (a_PI * (1 - e2_PI)) / Math.Sqrt(Math.Pow(1 - e2_PI * Math.Pow(Math.Sin(lat1Rad_PI), 2), 3))

        N2_PI = a_PI / Math.Sqrt(1 - e2_PI * Math.Pow(Math.Sin(lat2Rad_PI), 2))
        M2_PI = (a_PI * (1 - e2_PI)) / Math.Sqrt(Math.Pow(1 - e2_PI * Math.Pow(Math.Sin(lat2Rad_PI), 2), 3))

        N_m = (N1_PI + N2_PI) / 2
        M_m = (M1_PI + M2_PI) / 2

        latM = (lat1Rad_PI + lat2Rad_PI) / 2

        'Variacao da Latitude e Longitude'
        varLatPI = lat2Rad_PI - lat1Rad_PI
        varLongPI = long2Rad_PI - long1Rad_PI

        x = varLongPI * Math.Cos(latM) * N_m * (648000 / Math.PI)
        y = varLatPI * Math.Cos(varLongPI / 2) * M_m * 648000 / Math.PI

    End Sub

    Private Sub btnInputDados_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInputDados.Click
        txtSemiEixoMaiorPI.Text = 6378160
        txtInvAchatamentoPI.Text = 298.25
        txtLat1GPI.Text = 25
        txtLat1MPI.Text = 22
        txtLat1SPI.Text = 49.321
        txtLong1GPI.Text = 49
        txtLong1MPI.Text = 27
        txtLong1SPI.Text = 34.245

        txtLat2GPI.Text = 25
        txtLat2MPI.Text = 22
        txtLat2SPI.Text = 49.321
        txtLong2GPI.Text = 49
        txtLong2MPI.Text = 27
        txtLong2SPI.Text = 34.245
    End Sub

    Private Sub btnLimpar2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLimpar2.Click
        txtSemiEixoMaiorPI.Text = ""
        txtInvAchatamentoPI.Text = ""
        txtLat1GPI.Text = ""
        txtLat1MPI.Text = ""
        txtLat1SPI.Text = ""
        txtLong1GPI.Text = ""
        txtLong1MPI.Text = ""
        txtLong1SPI.Text = ""
        txtElementoF.Text = ""
        txtDifLat.Text = ""
        txtVarLat.Text = ""
        txtLat2GPI.Text = ""
        txtLat2MPI.Text = ""
        txtLat2SPI.Text = ""
        txtLong2GPI.Text = ""
        txtLong2MPI.Text = ""
        txtLong2SPI.Text = ""
        txtPrimExcPI.Text = ""
        txtLat1DecPI.Text = ""
        txtLat1RadPI.Text = ""
        txtLong1DecPI.Text = ""
        txtLong1RadPI.Text = ""
        txtLat2DecPI.Text = ""
        txtLat2RadPI.Text = ""
        txtLong2DecPI.Text = ""
        txtLong2RadPI.Text = ""
    End Sub
End Class
