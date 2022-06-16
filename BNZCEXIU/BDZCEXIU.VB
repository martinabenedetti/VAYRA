Imports System.IO
Imports NTSInformatica.CLN__STD
Imports NTSInformatica
Imports System.Data.SqlClient
Imports System.Data.OleDb


Public Class CLDZCEXIU
    Inherits CLD__BASE

    '    Public Overridable Function GetArtprox(ByVal strDitta As String, ByRef ds As DataSet) As Boolean
    '        Dim strSQL As String = ""
    '        Dim dttTmp As New DataTable
    '        Try

    '            strSQL = "
    'SELECT artprox.codditt, artprox.apx_codart, artico.ar_descr AS xxapx_codart, artico.ar_unmis AS xxar_unmis, 
    'CASE WHEN artprox.apx_esist < 0 THEN 0 ELSE artprox.apx_esist END AS apx_esist, 0.000 as xx_qtadainv, getdate() as xx_dtinv
    'FROM artprox
    'INNER JOIN artico
    '    ON artico.codditt = artprox.codditt
    '    AND artico.ar_codart = artprox.apx_codart
    'WHERE artprox.codditt = " & CStrSQL(strDitta)

    '            ' lista selezionata articoli
    '            If NTSCInt(ds.Tables("IMPOSTAZIONI").Rows(0)!ei_codlsar) <> 0 Then
    '                strSQL &= " 
    '                AND EXISTS (
    '                    SELECT TOP 1 lsa_codart 
    '                    FROM listsar 
    '                    WHERE 
    '                        listsar.codditt = artico.codditt 
    '                        AND listsar.lsa_codart = artico.ar_codart 
    '                        AND listsar.lsa_codlsar = " & CStrSQL(NTSCInt(ds.Tables("IMPOSTAZIONI").Rows(0)!ei_codlsar)) & "
    '                )"
    '            End If

    '            ' selezione
    '            If NTSCStr(ds.Tables("IMPOSTAZIONI").Rows(0)!ei_whereart).Trim.Length > 0 Then
    '                strSQL &= " AND ( 1 = 1 "
    '                TraduciWhere(NTSCStr(ds.Tables("IMPOSTAZIONI").Rows(0)!ei_whereart), strSQL)
    '                strSQL &= ")"
    '            End If

    '            strSQL &= " ORDER BY artico.ar_codart"

    '            ' leggo artprox
    '            If Not ds.Tables.Contains("ARTPROX") Then
    '                ds = OpenRecordset(strSQL, CLE__APP.DBTIPO.DBAZI, "ARTPROX", ds)
    '            Else
    '                ds.Tables("ARTPROX").Clear()
    '                dttTmp = OpenRecordset(strSQL, CLE__APP.DBTIPO.DBAZI)
    '                ds.Tables("ARTPROX").Merge(dttTmp)
    '            End If

    '            Return True

    '        Catch ex As Exception
    '            '--------------------------------------------------------------
    '            CLN__STD.GestErr(ex, Me, "")
    '            '--------------------------------------------------------------
    '        End Try
    '    End Function


    Public Overridable Function GetArtprox(ByVal strDitta As String, ByRef ds As DataSet) As Boolean
        Dim strSQL As String = ""
        Dim dttTmp As New DataTable
        Try

            strSQL = "
SELECT artpro.codditt, artpro.ap_codart, artico.ar_descr AS xxapx_codart, artico.ar_unmis AS xxar_unmis, 
CASE WHEN artpro.ap_esist < 0 THEN 0 ELSE artpro.ap_esist END AS ap_esist, 0.000 as xx_qtadainv, getdate() as xx_dtinv
FROM artpro
INNER JOIN artico
    ON artico.codditt = artpro.codditt
    AND artico.ar_codart = artpro.ap_codart
WHERE ap_magaz=1 AND artpro.codditt = " & CStrSQL(strDitta)

            ' lista selezionata articoli
            If NTSCInt(ds.Tables("IMPOSTAZIONI").Rows(0)!ei_codlsar) <> 0 Then
                strSQL &= " 
                AND EXISTS (
                    SELECT TOP 1 lsa_codart 
                    FROM listsar 
                    WHERE 
                        listsar.codditt = artico.codditt 
                        AND listsar.lsa_codart = artico.ar_codart 
                        AND listsar.lsa_codlsar = " & CStrSQL(NTSCInt(ds.Tables("IMPOSTAZIONI").Rows(0)!ei_codlsar)) & "
                )"
            End If

            ' selezione
            If NTSCStr(ds.Tables("IMPOSTAZIONI").Rows(0)!ei_whereart).Trim.Length > 0 Then
                strSQL &= " AND ( 1 = 1 "
                TraduciWhere(NTSCStr(ds.Tables("IMPOSTAZIONI").Rows(0)!ei_whereart), strSQL)
                strSQL &= ")"
            End If

            strSQL &= " ORDER BY artico.ar_codart"

            ' leggo artprox
            If Not ds.Tables.Contains("ARTPROX") Then
                ds = OpenRecordset(strSQL, CLE__APP.DBTIPO.DBAZI, "ARTPROX", ds)
            Else
                ds.Tables("ARTPROX").Clear()
                dttTmp = OpenRecordset(strSQL, CLE__APP.DBTIPO.DBAZI)
                ds.Tables("ARTPROX").Merge(dttTmp)
            End If

            Return True

        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try
    End Function


End Class
