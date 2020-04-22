<%@ Page Language="vb" AutoEventWireup="false"%>

<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.IO" %>
<script runat="server">



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim start_time As DateTime = Now

        Dim mukey As String = Request.QueryString("mukey")
        Dim outtab1 As String
        Dim outtab3 As String
        Dim outtab4 As String

        Dim muSQLstring As String = "SELECT lkey,musym,muacres,mukey,muname,farmlndcl,muhelcl,muwathelcl, muwndhelcl, interpfocus, mukind, mustatus FROM soils2019.mapunit where mukey = '" & mukey & "';"
        Dim mudatatable As DataTable = getSqlDataTable(muSQLstring)
        Dim mudata As DataRow = mudatatable.Rows(0)

        Dim maSQL As String = "SELECT * FROM soils2019.muaggatt where mukey = '" & mukey & "';"
        Dim madatatab As DataTable = getSqlDataTable(maSQL)
        Dim madata As DataRow = madatatab.Rows(0)

        Dim legSQL As String = "SELECT * FROM soils2019.legend where lkey = '" & mudata("lkey") & "' ;"
        Dim Ldatatab As DataTable = getSqlDataTable(legSQL)
        Dim Ldata As DataRow = Ldatatab.Rows(0)
        Dim c1 As String = Ldata("areaname")
        Dim c2 As Array = Split(c1, ",")
        Dim state As String
        If UBound(c2) > 0 Then
            state = c2(1)
        Else
            state = "N/A"
        End If
        Dim county As String = c2(0)
        Dim symbol As String = Ldata("areasymbol")

        Dim cSQL As String = "SELECT * FROM soils2019.component where mukey = '" & mukey & "' and majcompflag = 'Yes';"
        Dim cdatatab As DataTable = getSqlDataTable(cSQL)

        Dim pc1 As String() = Split(mudata("muname"), " ")
        Dim pc2 As String() = Split(pc1(0), "-")

        'outtab1 = "<table style='width: 100%'>"
        outtab1 = "<div style='border:solid 1px #000000;'>"
        outtab1 += "<div style='background-color:#0099CC;color:#ffffff;font-size:26px;font-weight:bold;padding:20px;'>Soil Data Map Unit Interpretation Report</div>"

        Dim muacres As String = "N/A"
        If Not IsDBNull(mudata("muacres")) Then
            muacres = FormatNumber(mudata("muacres"), 0)
        End If
        outtab1 += "<div style='padding:15px;font-size:11pt;'>"
        outtab1 += "<div class='report_header'>Major Map Unit Introduction</div>"
        outtab1 += "<div class='report_content'>"
        outtab1 += "<div class='report_subheader'>" & mudata("musym") & " - " & mudata("muname") & "</div>"


        outtab1 += "<div class='inner_content'>The spatial extent of this map unit is " & muacres & " acres "
        If Not IsDBNull(mudata("muacres")) Then
            outtab1 += "(" & FormatNumber(mudata("muacres") / 640, 1) & " square miles), comprising "
            outtab1 += FormatNumber(100 * mudata("muacres") / Ldata("areaacres"), 1) & "% of the total survey area of " & c1 & "." & vbCrLf
        Else
            outtab1 += "in " & c1 & vbCrLf
        End If
        If Not IsDBNull(mudata("muwndhelcl")) Then 'HIGHLY ERODIBLE LAND AND PRIMEFARMLAND RATING
            Dim wathel As Array
            If IsDBNull(mudata("muwathelcl")) Then
                wathel = Split(mudata("muwndhelcl"), " land")
            Else
                wathel = Split(mudata("muwathelcl"), " land")
            End If
            Dim wndhel As Array = Split(mudata("muwndhelcl"), " land")
            outtab1 += "&nbsp; This map unit is " & LCase(wathel(0)) & " by water; " & LCase(wndhel(0)) & " by wind." & vbCrLf
            If InStr(mudata("farmlndcl"), "All") <> 0 Then
                outtab1 += "&nbsp; This map unit is " & LCase(mudata("farmlndcl")) & "." & vbCrLf
            Else
                outtab1 += "&nbsp; " & mudata("farmlndcl") & "." & vbCrLf
            End If
        End If
        outtab1 += "</div>" & vbCrLf

        If Not IsDBNull(madata("hydgrpdcd")) Then 'HIGHLY ERODIBLE LAND AND PRIMEFARMLAND RATING
            outtab1 += "<div class='report_subheader'>Map Unit Characteristics</div>"
            outtab1 += "<div class='inner_content'><ul style='position:relative;top:-15px;'>"
            'outtab1 += "<TR><TD colspan='2' class='h3'>Map Unit Characteristics</TD>" & vbCrLf

            If IsDBNull(madata("drclassdcd")) Then
                outtab1 += "<li>This map unit "
            Else
                outtab1 += "<li>This " & LCase(madata("drclassdcd")) & " map unit "
            End If

            outtab1 += "is " & LCase(madata("hydclprs")) & "% hydric and is classed in Hydrologic Group " & madata("hydgrpdcd") & ".</li>" & vbCrLf
            If IsDBNull(madata("wtdepaprjunmin")) Then
                outtab1 += "<li>No seasonal water table is present.</li>" & vbCrLf
            Else
                outtab1 += "<li>A seasonal water table is present at " & FormatNumber(madata("wtdepaprjunmin") * 0.39, 1) & " inches.</li>" & vbCrLf
            End If
            outtab1 += "<li>Ponding is present " & madata("pondfreqprs") & "% of the time.</li>" & vbCrLf
            If IsDBNull(madata("flodfreqdcd")) Then
                outtab1 += "<li>Flooding frequency is not rated for this map unit.</li>" & vbCrLf
            Else
                If InStr(madata("flodfreqdcd"), "None") <> 0 Then
                    outtab1 += "<li>This is not a flood plain unit.</li>" & vbCrLf
                Else
                    outtab1 += "<li>This land is subject to " & LCase(madata("flodfreqdcd")) & " flooding.</li>" & vbCrLf
                End If
            End If

            If Not IsDBNull(madata("slopegradwta")) Then
                outtab1 += "<li>The average slope gradient is " & FormatNumber(madata("slopegradwta"), 1) & "%.</li>" & vbCrLf
            End If

            ' response.write "&nbsp; The non-irrigated capability class rating is "&MAdata("niccdcd")& ".<BR>" & vbcrlf


            'View Distribution Map
            'If InStr(state, "Missouri") > 0 Then
            'outtab1 += c1
            'If Left(c1, 8) = "St Louis" Then
            'outtab1 += "<li><a target='_blank' href='../distribution_2006/dist_" & Left(UCase(pc2(0)), 1) & "/" & UCase(pc2(0)) & "_St_Louis.html'>"
            'outtab1 += "View Map Unit Distribution</A></li>" & vbCrLf
            'Else
            'outtab1 += "<li><a target='_blank' href='../distribution_2006/dist_" & Left(UCase(pc2(0)), 1) & "/" & UCase(pc2(0)) & "_" & Replace(county, " County", "")
            'outtab1 += ".html'>View Map Unit Distribution</A></li>" & vbCrLf
            'End If
            'End If
            outtab1 += "</ul></div>" & vbCrLf
            'outtab1 += "<div class='report_subheader'>Avail Water Cap (in)</div>" & vbCrLf
            'outtab1 += "<div class='inner_content'><table><tr><td style='font-weight:bold;'>Depth</td><td style='font-weight:bold;'>Water</td></tr>" & vbCrLf


            'outtab1 += "<TR><TD class='dataleftsmall' style='height:50%'>0-10 in</TD>" & vbCrLf

            If Not IsDBNull(madata("aws025wta")) Then
                'outtab1 += "    <TD class='dataleftsmall' style='height:50%'>" & FormatNumber(madata("aws025wta") / 2.54, 1) & "</TD></TR>" & vbCrLf
            End If
            'outtab1 += "<TR><TD class='dataleftsmall' style='height:50%'>0-20 in</TD>" & vbCrLf
            If Not IsDBNull(madata("aws050wta")) Then
                'outtab1 += "    <TD class='dataleftsmall' style='height:50%'>" & FormatNumber(madata("aws050wta") / 2.54, 1) & "</TD></TR>" & vbCrLf
            End If
            ' outtab1 += "<TR><TD class='dataleftsmall' style='height:50%'>0-40 in</TD>" & vbCrLf
            If Not IsDBNull(madata("aws0100wta")) Then
                'outtab1 += "    <TD class='dataleftsmall' style='height:50%'>" & FormatNumber(madata("aws0100wta") / 2.54, 1) & "</TD></TR>" & vbCrLf
            End If
            'outtab1 += "<TR><TD class='dataleftsmall' style='height:50%'>0-60 in</TD>" & vbCrLf
            If Not IsDBNull(madata("aws0150wta")) Then
                'outtab1 += "    <TD class='dataleftsmall' style='height:50%'>" & FormatNumber(madata("aws0150wta") / 2.54, 1) & "</TD></TR>" & vbCrLf
            End If
            'outtab1 += "</TR>" & vbCrLf
        End If


        'outtab1 += "</table></div></div>"
        'top.Text = outtab1

        'prepare component descriptions
        Dim outtab2 As String


        Dim mtx1SQL As String
        Dim mtx1datatable As DataTable

        Dim cmpSQL As String
        Dim cmpdatatable As DataTable

        'outtab1 += "<div class='report_header'>Major Map Unit Components</div>"

        If cdatatab.Rows.Count = 1 Then
            outtab1 += "<input type='hidden' class='mult_comps' value='0' />"
            outtab1 += "<div class='report_header'>Major Map Unit Component</div>" & vbCrLf
        Else
            outtab1 += "<input type='hidden' class='mult_comps' value='1' />"
            outtab1 += "<div class='report_header'>Major Map Unit Components</div>" & vbCrLf
        End If

        outtab1 += "<div class='report_content'>"
        outtab1 += "<div class='accordion'><span style='font-style:italic'>Click to expand/contract each section.</span>"

        For Each cdata As DataRow In cdatatab.Rows
            Dim localphase As String = ""
            If Not IsDBNull(cdata("localphase")) Then
                If cdata("localphase") <> "" Then
                    localphase = " (" & cdata("localphase") & ")"
                End If
            End If
            'Component name, kind and classification
            outtab1 += "<div class='accordion-toggle'>" & UCase(Left(cdata("compname"), 1)) & LCase(Right(cdata("compname"), Len(cdata("compname")) - 1)) & " " & cdata("compkind") & localphase & "</B> - <span style='font-size:14px; font-weight:normal;'>" & cdata("taxclname") & "</div>"
            outtab1 += "<div class='accordion-content'>"
            outtab1 += "<div class='report_subheader'>Basic Information</div>"
            outtab1 += "<table><TR><TD colspan='4' style='text-align:left;'><B>" & UCase(Left(cdata("compname"), 1)) & LCase(Right(cdata("compname"), Len(cdata("compname")) - 1)) & " "
            outtab1 += cdata("compkind") & localphase & "</B> - <span style='font-size:14px; font-weight:normal;'>" & cdata("taxclname") & "" & vbCrLf

            mtx1SQL = "SELECT * FROM soils2019.mutext where mukey = '" & mukey & "' and textcat = 'SOIL' and text like '%" & cdata("compname") & "%' ;"
            mtx1datatable = getSqlDataTable(mtx1SQL)

            If mtx1datatable.Rows.Count = 1 Then
                Dim nontechsoil As String = Replace(mtx1datatable.Rows(0)("text"), mudata("musym") & " " & mudata("muname"), "")
                outtab1 += "<BR>&nbsp;&nbsp;&nbsp;&nbsp;" & nontechsoil & " "
            Else
                cmpSQL = "SELECT * FROM soils2019.copmgrp where cokey = '" & cdata("cokey") & "';"
                cmpdatatable = getSqlDataTable(cmpSQL)
                If cmpdatatable.Rows.Count > 0 Then
                    outtab1 += "<BR>These soils are formed in " & cmpdatatable.Rows(0)("pmgroupname")
                Else
                    outtab1 += "<BR>These soils are "
                End If
                If Not IsDBNull(cdata("geomdesc")) Then
                    If StrComp(Right(cdata("geomdesc"), 1), "s") = 0 Then
                        outtab1 += " located on " & cdata("geomdesc")
                    Else
                        outtab1 += " located on " & cdata("geomdesc") & "s"
                    End If
                End If
                If Not IsDBNull(cdata("earthcovkind1")) Then
                    outtab1 += "  under " & LCase(cdata("earthcovkind1"))
                End If
                If Not IsDBNull(cdata("earthcovkind2")) Then
                    outtab1 += " and " & LCase(cdata("earthcovkind2"))
                End If
                outtab1 += ", and comprise approximately " & cdata("comppct_r") & "% of the map unit. "
                If Not IsDBNull(cdata("runoff")) Then
                    outtab1 += " The surface water runoff class is " & LCase(cdata("runoff"))
                End If
                If Not IsDBNull(cdata("drainagecl")) Then
                    outtab1 += " and the natural drainage condition of the soil is " & LCase(cdata("drainagecl")) & ". "
                Else
                    outtab1 += "."
                End If

                If IsNumeric(madata("wtdepannmin")) Then
                    outtab1 += "The top of the seasonal high water table is at " & FormatNumber(madata("wtdepannmin") * 0.39, 0) & " inches."
                Else
                    outtab1 += "The seasonal high water table is at a depth of more than 6 feet."
                End If
                If IsNumeric(cdata("nirrcapcl")) Then
                    outtab1 += " This map unit component is assigned to the nonirrigated land capability classification " & cdata("nirrcapcl")
                    If Not IsDBNull(cdata("nirrcapscl")) Then
                        outtab1 += LCase(cdata("nirrcapscl")) & ". "
                    Else
                        outtab1 += ". "
                    End If
                Else
                    outtab1 += " This map unit component is not assigned a land capability classification. "
                End If

                outtab1 += "</span></TD></TR>" & vbCrLf
            End If


            'outtab1 += "<TR class='comptabrow2'><TD>"

            If Not IsDBNull(cdata("map_l")) Or Not IsDBNull(cdata("map_h")) Then
                outtab1 += "<TR class='comptabrow2'><TD class='rjustify smallft'><B>Annual Precip:</B></td><td class='ljustify'>" & FormatNumber(cdata("map_l") / 25.4, 0) & " - "
                outtab1 += FormatNumber(cdata("map_h") / 25.4, 0) & " inches</TD>" & vbCrLf
            Else
                outtab1 += "<TR class='comptabrow2'><TD class='rjustify smallft'><B>Annual Precip:</B></td><td class='ljustify'>No Data Available</TD>" & vbCrLf
            End If

            outtab1 += "    <TD class='rjustify smallft'><B>Annual Air Temp:</B></td><td class='ljustify'>"
            If IsNumeric(cdata("airtempa_l")) Then
                outtab1 += FormatNumber((32 + cdata("airtempa_l") * 1.8), 0) & " - " & FormatNumber((32 + cdata("airtempa_h") * 1.8), 0) & " <sup>o</sup>F ("
                outtab1 += FormatNumber(cdata("airtempa_l"), 0) & " - " & FormatNumber(cdata("airtempa_h"), 0) & "<sup>o</sup>C) </TD></TR>" & vbCrLf
            Else
                outtab1 += " N/A</TD></TR>" & vbCrLf
            End If
            outtab1 += "<TR class='comptabrow2'><TD class='rjustify smallft'><B>Frost Free:</B></td><td class='ljustify'>"
            If IsNumeric(cdata("airtempa_l")) Then
                outtab1 += cdata("ffd_l") & " - " & cdata("ffd_h") & " days</TD>" & vbCrLf
            Else
                outtab1 += " N/A</TD>" & vbCrLf
            End If
            outtab1 += "<TD class='rjustify smallft'><B>Dry Albedo (rep):</B></td><td class='ljustify'>"
            If IsNumeric(cdata("albedodry_r")) Then
                outtab1 += FormatNumber(cdata("albedodry_r"), 2) & "</TD></TR>" & vbCrLf
            Else
                outtab1 += " N/A</TD></TR>" & vbCrLf
            End If

            If IsNumeric(cdata("map_r")) And IsNumeric(cdata("airtempa_r")) Then
                outtab1 += "<TR class='comptabrow2'><TD class='rjustify smallft'><B>Precipation Effectiveness<span style='color:red'>*</span>:</B></td><td class='ljustify'>"
                outtab1 += FormatNumber(87 * (((FormatNumber(cdata("map_r") / 25.4, 1)) / ((FormatNumber((32 + cdata("airtempa_r") * 1.8), 1)) - 10)) ^ 1.11111), 2) & "</TD>" & vbCrLf
            Else
                outtab1 += "<TR class='comptabrow2'><TD class='rjustify smallft'><B>Precipation Effectiveness:</B></td><td class='ljustify'>No Data Available</TD>" & vbCrLf
            End If
            If IsNumeric(cdata("map_r")) And IsNumeric(cdata("airtempa_r")) Then
                outtab1 += "<TD class='rjustify smallft'><B>Climate Factor<span style='color:red'>**</span>:</B></td><td class='ljustify'>"
                outtab1 += FormatNumber(34.48 * 8 ^ 3 / (87 * (((FormatNumber(cdata("map_r") / 25.4, 1)) / ((FormatNumber((32 + cdata("airtempa_r") * 1.8), 1)) - 10)) ^ 1.11111)) ^ 2, 2) & "</TD></TR>" & vbCrLf
            Else
                outtab1 += "<TD class='rjustify smallft'><B>Climate Factor:</B></td><td class='ljustify'>No Data Available</TD></TR>" & vbCrLf
            End If
            outtab1 += "<TR class='comptabrow2'> <TD class='rjustify smallft'><B>Erosion Class:</B></td><td class='ljustify'>" & cdata("erocl") & "</TD>" & vbCrLf
            outtab1 += "     <TD class='rjustify smallft'><B>Runoff Class:</B></td><td class='ljustify'>" & cdata("runoff") & "</TD></TR>" & vbCrLf
            If IsNumeric(cdata("slope_l")) And IsNumeric(cdata("slope_h")) Then
                outtab1 += "<TR class='comptabrow2'> <TD class='rjustify smallft'><B>Slope Gradient:</B></td><td class='ljustify'>" & FormatNumber(cdata("slope_l"), 0) & " - " & FormatNumber(cdata("slope_h"), 0) & "%</TD>" & vbCrLf
            Else
                outtab1 += "<TR class='comptabrow2'> <TD class='rjustify smallft'><B>Slope Gradient:</B></td><td class='ljustify'>N/A</TD>" & vbCrLf
            End If
            If IsNumeric(cdata("slopelenusle_r")) Then
                outtab1 += " <TD class='rjustify smallft'><B>USLE Slope Length (rep):</B></td><td class='ljustify'>" & FormatNumber(cdata("slopelenusle_r") * 3.2808, 0) & " ft (" & cdata("slopelenusle_r") & " M)</TD></TR>" & vbCrLf
            Else
                outtab1 += " <TD class='rjustify smallft'><B>USLE Slope Length (rep):</B></td><td class='ljustify'>N/A</TD></TR>" & vbCrLf
            End If
            outtab1 += "<TR class='comptabrow2'> <TD class='rjustify smallft'><B>USLE T Factor:</B></td><td class='ljustify'>" & cdata("tfact") & " tons/acre-year</TD>" & vbCrLf
            outtab1 += "     <TD class='rjustify smallft'><B>Drainage Class:</B></td><td class='ljustify'>" & cdata("drainagecl") & "</TD></TR>" & vbCrLf
            outtab1 += "<TR class='comptabrow2'> <TD class='rjustify smallft'><B>Wind Erodibility Group:</B></td><td class='ljustify'>" & cdata("weg") & "</TD>" & vbCrLf
            outtab1 += "     <TD class='rjustify smallft'><B>Wind Erodibility Index:</B></td><td class='ljustify'>" & cdata("wei") & "</TD></TR>" & vbCrLf
            outtab1 += "<TR class='comptabrow2'> <TD class='rjustify smallft'><B>Hydric:</B></td><td class='ljustify'>" & cdata("hydricrating") & vbCrLf
            If Not IsDBNull(cdata("hydricrating")) Then
                If InStr(cdata("hydricrating"), "Yes") <> 0 Then
                    outtab1 += " - " & cdata("hydricon") & "</FONT>" & vbCrLf
                End If
            End If

            outtab1 += "<TD class='rjustify smallft'><B>Land Capability Class:</B></td><td class='ljustify'>" & cdata("nirrcapcl")
            If Not IsDBNull(cdata("nirrcapscl")) Then
                outtab1 += LCase(cdata("nirrcapscl"))
            End If
            outtab1 += "</TD></TR>" & vbCrLf
            outtab1 += "<TR class='comptabrow2'><TD colspan='4' class='dataleftsmall ljustify'><span style='color:red'>*</span> Thorntwaite Precipitation-Effectiveness index: "
            outtab1 += "PE = 87 * (annual average precipitation<sub>in</sub> * annual average temperature<sub><sup>o</sup>F</sub>)<sup>1.1111</sup><BR>"
            outtab1 += "<span style='color:red'>**</span> Climate Factor: C = 34.48 * V<sup>3</sup> / PE<sup>2</sup> -- V = average annual wind velocity (8 mph used here)</TD></TR>" & vbCrLf

            outtab1 += "<TR class='comptabrow2'><TD colspan='4' class='ljustify'><BR>View Series Distribution: " & vbCrLf

            If InStr(cdata("compkind"), "Series") > 0 Or InStr(cdata("compkind"), "Variant") > 0 Or InStr(cdata("compkind"), "Taxadjunct") > 0 Then
                Dim CARESMapName As String
                'If InStr(state, "Missouri") > 0 Then
                'outtab2 += "&nbsp;&nbsp;<a target='_blank' href='../distribution_2006/dist_" & Left(UCase(pc2(0)), 1) & "/" & UCase(pc2(0)) & ".html'>CSS Site</A>" & vbCrLf
                'End If
                CARESMapName = Replace(cdata("compname"), " ", "_") + "_" + cdata("compkind")
                CARESMapName = Replace(CARESMapName, ".", "_")
                CARESMapName = Replace(CARESMapName, ",", "_")
                CARESMapName = Replace(CARESMapName, "(", "_")
                CARESMapName = Replace(CARESMapName, ")", "_")
                CARESMapName = Replace(CARESMapName, "-", "_")
                CARESMapName = Replace(CARESMapName, "'", "_")

                outtab1 += "&nbsp;&nbsp;<a target='_blank' href='/_data/images/soils_distribution/" & CARESMapName & ".jpg'>CARES Maps</A>" & vbCrLf
                'outtab1 += "&nbsp;&nbsp;<a target='_blank' href='http://www.cei.psu.edu/soiltool/semtool.html?seriesname=" & UCase(cdata("compname")) & "'> "
                'outtab1 += "SEMTOOL</A>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; " & vbCrLf
                outtab1 += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a target='_blank' "
                outtab1 += "href='https://soilseries.sc.egov.usda.gov/OSD_Docs/" & Left(cdata("compname"), 1) & "/" & UCase(cdata("compname")) & ".html'>Official Soil Series Description (NRCS)</A>" & vbCrLf
            Else
                outtab1 += "&nbsp;&nbsp;No distribution available at this time"
            End If
            outtab1 += "</TD></TR></table><br />" & vbCrLf
            outtab1 += "<div class='report_subheader'>Horizon Data</div>"
            'outtab2 += "</TD></TR>" & vbCrLf

            Dim chrSQL As String = "SELECT * FROM soils2019.chorizon where cokey = '" & cdata("cokey") & "' order by hzdept_r ASC ;"
            Dim chrdatatab As DataTable = getSqlDataTable(chrSQL)

            If chrdatatab.Rows.Count > 0 Then
                outtab1 += "<table>" & vbCrLf
                ' outtab1 += "<TR class='comptabrow2'><TD colspan='4'><TABLE bgcolor='lightcyan' width='100%' border='0' cellpadding='0' cellspacing='0' align='right'>" & vbCrLf
                outtab1 += "<TR class='comptabrow1'><TD colspan='1'>&nbsp;</td> <TD colspan='14' class='dataleft'><B>Horizon Data: Key Physical and Chemical Properties</B> (representative values)</TD></TR>" & vbCrLf
                outtab1 += "<TR class='comptabrow1'>"
                outtab1 += "<TH colspan='1' width='1%' class='ljustify'>&nbsp;</TD>" & vbCrLf
                outtab1 += "<TH colspan='1' width='4%' class='ljustify'>Name </TD>" & vbCrLf
                outtab1 += "<TH colspan='1' width='8%' class='ljustify'>Depth</TD>" & vbCrLf
                outtab1 += "<TH colspan='1' width='9%' class='ljustify'>Clay</TD>" & vbCrLf
                outtab1 += "<TH colspan='1' width='9%' class='ljustify'>Silt</TD>" & vbCrLf
                outtab1 += "<TH colspan='1' width='9%' class='ljustify'>Sand</TD>" & vbCrLf
                outtab1 += "<TD colspan='1' width='9%' class='ljustify'><B>Frags</B><BR>2mm-3&#148;</TD>" & vbCrLf
                outtab1 += "<TD colspan='1' width='9%' class='ljustify'><B>Frags</B><BR>3&#148;-10&#148;</TD>" & vbCrLf
                outtab1 += "<TD colspan='1' width='9%' class='ljustify'><B>Frags</B><BR>>10&#148;</TD>" & vbCrLf
                outtab1 += "<TD colspan='1' width='9%' class='ljustify'><B>OM</B></TD>" & vbCrLf
                outtab1 += "<TD colspan='1' width='9%' class='ljustify'><B>CaCO<sub>3</sub></B></TD>" & vbCrLf
                outtab1 += "<TD colspan='1' width='9%' class='ljustify'><B>pH</B> <font size='1'> (H<sub>2</sub>O)</font></TD>" & vbCrLf
                outtab1 += "<TD colspan='1' width='9%' class='ljustify'><B>D<sub>b</sub></B><font size='1'>(Dry)</font></TD>" & vbCrLf
                outtab1 += "<TD colspan='1' width='8%' class='ljustify'><B>CEC-7</B></TD>" & vbCrLf
                outtab1 += "<TD colspan='1' width='2%' class='ljustify'>&nbsp;</TD>" & vbCrLf


                For Each chrdata In chrdatatab.Rows
                    outtab1 += "<TR><TD colspan='1' width='1%' class='datacenter'>&nbsp;</TD>" & vbCrLf

                    outtab1 += "<TD class='dataleft'><B>" & chrdata("hzname") & "</B></TD>"
                    outtab1 += "<TD class='datacenter'>" & vbCrLf
                    If Not IsDBNull(chrdata("hzdept_r")) Then
                        outtab1 += FormatNumber(chrdata("hzdept_r") / 2.54, 0)
                    End If
                    If Not IsDBNull(chrdata("hzdepb_r")) Then
                        outtab1 += " - " & FormatNumber(chrdata("hzdepb_r") / 2.54, 0) & vbCrLf
                    End If
                    outtab1 += " in</TD>" & vbCrLf
                    If Not IsDBNull(chrdata("claytotal_r")) Then
                        outtab1 += "<TD class='datacenter'>" & FormatNumber(chrdata("claytotal_r"), 1) & "%</TD>" & vbCrLf
                    Else
                        outtab1 += "<TD class='datacenter'>N/A</TD>" & vbCrLf
                    End If
                    If Not IsDBNull(chrdata("silttotal_r")) Then
                        outtab1 += "<TD class='datacenter'>" & FormatNumber(chrdata("silttotal_r"), 1) & "%</TD>" & vbCrLf
                    Else
                        outtab1 += "<TD class='datacenter'>N/A</TD>" & vbCrLf
                    End If
                    If Not IsDBNull(chrdata("sandtotal_r")) Then
                        outtab1 += "<TD class='datacenter'>" & FormatNumber(chrdata("sandtotal_r"), 1) & "%</TD>" & vbCrLf
                    Else
                        outtab1 += "<TD class='datacenter'>N/A</TD>" & vbCrLf
                    End If
                    If IsNumeric(chrdata("sieveno10_r")) Then
                        outtab1 += "<TD class='datacenter'>" & FormatNumber(100 - chrdata("sieveno10_r"), 0) & "%</TD>" & vbCrLf
                    Else
                        outtab1 += "<TD class='datacenter'>N/A</TD>" & vbCrLf
                    End If
                    If IsNumeric(chrdata("frag3to10_r")) Then
                        outtab1 += "<TD class='datacenter'>" & FormatNumber(chrdata("frag3to10_r"), 0) & "%</TD>" & vbCrLf
                    Else
                        outtab1 += "<TD class='datacenter'>N/A</TD>" & vbCrLf
                    End If
                    If Not IsDBNull(chrdata("fraggt10_r")) Then
                        outtab1 += "<TD class='datacenter'>" & FormatNumber(chrdata("fraggt10_r"), 0) & "%</TD>" & vbCrLf
                    Else
                        outtab1 += "<TD class='datacenter'>N/A</TD>" & vbCrLf
                    End If
                    If Not IsDBNull(chrdata("om_r")) Then
                        outtab1 += "<TD class='datacenter'>" & FormatNumber(chrdata("om_r"), 1) & "%</TD>" & vbCrLf
                    Else
                        outtab1 += "<TD class='datacenter'>N/A</TD>" & vbCrLf
                    End If
                    If Not IsDBNull(chrdata("caco3_r")) Then
                        outtab1 += "<TD class='datacenter'>" & chrdata("caco3_r") & "%</TD>" & vbCrLf
                    Else
                        outtab1 += "<TD class='datacenter'>N/A</TD>" & vbCrLf
                    End If
                    If Not IsDBNull(chrdata("ph1to1h2o_r")) Then
                        outtab1 += "<TD class='datacenter'>" & FormatNumber(chrdata("ph1to1h2o_r"), 1) & "</TD>" & vbCrLf
                    Else
                        outtab1 += "<TD class='datacenter'>N/A</TD>" & vbCrLf
                    End If
                    If Not IsDBNull(chrdata("dbovendry_r")) Then
                        outtab1 += "<TD class='datacenter'>" & FormatNumber(chrdata("dbovendry_r"), 2) & "<font size='1'>g/cm<sup>3</sup></TD>" & vbCrLf
                    Else
                        outtab1 += "<TD class='datacenter'>N/A</TD>" & vbCrLf
                    End If
                    If Not IsDBNull(chrdata("cec7_r")) Then
                        outtab1 += "<TD class='datacenter'>" & FormatNumber(chrdata("cec7_r"), 1) & "</TD>" & vbCrLf
                    Else
                        outtab1 += "<TD class='datacenter'>N/A</TD>" & vbCrLf
                    End If
                    outtab1 += "<TD class='datacenter'>&nbsp;</TD>" & vbCrLf
                Next
                'outtab1 += "</TABLE></TD></TR>" & vbCrLf

                'outtab1 += "<TR class='comptabrow2'><TD colspan='4' class='dataleft'>&nbsp;</FONT></TD></TR>" & vbCrLf
                'outtab1 += "<TR class='comptabrow2'><TD colspan='4'><TABLE bgcolor='lightcyan' width='100%' border='0' cellpadding='0' cellspacing='0' align='right'>" & vbCrLf
                outtab1 += "<TR class='comptabrow1'><TD colspan='1'>&nbsp;</td> <TD colspan='14' class='dataleft'><B>Horizon Data: Moisture Characteristics</B>  (representative values)</TD></TR>" & vbCrLf
                outtab1 += "<TR class='comptabrow1'>"
                outtab1 += "<TD colspan='1' width='1%' class='datacenter'>&nbsp;</TD>" & vbCrLf
                outtab1 += "<TD colspan='1' width='4%' class='ljustify'><B>Name</B> </TD>" & vbCrLf
                outtab1 += "<TD colspan='1' width='8%' class='ljustify'><B>Depth</B></TD>" & vbCrLf
                outtab1 += "<TD colspan='1' width='8%' class='ljustify'><B>K<sub>sat</sub></B></TD>" & vbCrLf
                outtab1 += "<TD colspan='1' width='9%' class='ljustify'><B>AWC</B></TD>" & vbCrLf
                outtab1 += "<TD colspan='1' width='8%' class='ljustify'><B>H<sub>2</sub>O</B><BR>sat</TD>" & vbCrLf
                outtab1 += "<TD colspan='1' width='8%' class='ljustify'><B>H<sub>2</sub>O</B><BR>.33 bar</TD>" & vbCrLf
                outtab1 += "<TD colspan='1' width='8%' class='ljustify'><B>H<sub>2</sub>O</B><BR>15 bar</TD>" & vbCrLf
                outtab1 += "<TD colspan='1' width='8%' class='ljustify'><B>D<sub>b</sub></B><BR>.1 bar</TD>" & vbCrLf
                outtab1 += "<TD colspan='1' width='8%' class='ljustify'><B>D<sub>b</sub></B><BR>.33 bar</TD>" & vbCrLf
                outtab1 += "<TD colspan='1' width='8%' class='ljustify'><B>D<sub>b</sub></B><BR>15 bar</TD>" & vbCrLf
                outtab1 += "<TD colspan='1' width='8%' class='ljustify'><B>LEP</B></TD>" & vbCrLf
                outtab1 += "<TD colspan='1' width='7%' class='ljustify'><B>Liquid Limit</B></TD>" & vbCrLf
                outtab1 += "<TD colspan='1' width='7%' class='ljustify'><B>Plast Index</B></TD>" & vbCrLf
                outtab1 += "<TD colspan='1' width='2%' class='datacenter'>&nbsp;</TD>" & vbCrLf

                For Each chrdata In chrdatatab.Rows
                    outtab1 += "<TR>"
                    outtab1 += "<TD colspan='1' width='1%' class='datacenter'>&nbsp;</TD>" & vbCrLf
                    outtab1 += "<TD class='dataleft'><B>" & chrdata("hzname") & "</B></TD>" & vbCrLf
                    outtab1 += "<TD class='datacenter'>" & vbCrLf
                    If Not IsDBNull(chrdata("hzdept_r")) Then
                        outtab1 += FormatNumber(chrdata("hzdept_r") / 2.54, 0)
                    End If
                    If Not IsDBNull(chrdata("hzdepb_r")) Then
                        outtab1 += " - " & FormatNumber(chrdata("hzdepb_r") / 2.54, 0) & vbCrLf
                    End If
                    outtab1 += " in</TD>" & vbCrLf
                    If Not IsDBNull(chrdata("ksat_r")) Then
                        outtab1 += "<TD class='datacenter'>" & FormatNumber(chrdata("ksat_r") * 0.1417, 2) & "<font size='1'>in/hr</TD>" & vbCrLf
                    Else
                        outtab1 += "<TD class='datacenter'>N/A</TD>" & vbCrLf
                    End If
                    If Not IsDBNull(chrdata("awc_r")) Then
                        outtab1 += "<TD class='datacenter'>" & FormatNumber(chrdata("awc_r"), 2) & "<font size='1'>in/in</TD>" & vbCrLf
                    Else
                        outtab1 += "<TD class='datacenter'>N/A</TD>" & vbCrLf
                    End If
                    If Not IsDBNull(chrdata("wsatiated_r")) Then
                        outtab1 += "<TD class='datacenter'>" & chrdata("wsatiated_r") & "%</TD>" & vbCrLf
                    Else
                        outtab1 += "<TD class='datacenter'>N/A</TD>" & vbCrLf
                    End If
                    If Not IsDBNull(chrdata("wthirdbar_r")) Then
                        outtab1 += "<TD class='datacenter'>" & FormatNumber(chrdata("wthirdbar_r"), 1) & "%</TD>" & vbCrLf
                    Else
                        outtab1 += "<TD class='datacenter'>N/A</TD>" & vbCrLf
                    End If
                    If Not IsDBNull(chrdata("wfifteenbar_r")) Then
                        outtab1 += "<TD class='datacenter'>" & FormatNumber(chrdata("wfifteenbar_r"), 1) & "%</TD>" & vbCrLf
                    Else
                        outtab1 += "<TD class='datacenter'>N/A</TD>" & vbCrLf
                    End If

                    If Not IsDBNull(chrdata("dbtenthbar_r")) Then
                        outtab1 += "<TD class='datacenter'>" & FormatNumber(chrdata("dbtenthbar_r"), 2) & "<font size='1'>g/cm<sup>3</sup></TD>" & vbCrLf
                    Else
                        outtab1 += "<TD class='datacenter'>N/A</TD>" & vbCrLf
                    End If
                    If Not IsDBNull(chrdata("dbthirdbar_r")) Then
                        outtab1 += "<TD class='datacenter'>" & FormatNumber(chrdata("dbthirdbar_r"), 2) & "<font size='1'>g/cm<sup>3</sup></TD>" & vbCrLf
                    Else
                        outtab1 += "<TD class='datacenter'>N/A</TD>" & vbCrLf
                    End If
                    If Not IsDBNull(chrdata("dbfifteenbar_r")) Then
                        outtab1 += "<TD class='datacenter'>" & FormatNumber(chrdata("dbfifteenbar_r"), 2) & "<font size='1'>g/cm<sup>3</sup></TD>" & vbCrLf
                    Else
                        outtab1 += "<TD class='datacenter'>N/A</TD>" & vbCrLf
                    End If
                    If Not IsDBNull(chrdata("lep_r")) Then
                        outtab1 += "<TD class='datacenter'>" & FormatNumber(chrdata("lep_r"), 1) & "</TD>" & vbCrLf
                    Else
                        outtab1 += "<TD class='datacenter'>N/A</TD>" & vbCrLf
                    End If
                    If Not IsDBNull(chrdata("ll_r")) Then
                        outtab1 += "<TD class='datacenter'>" & FormatNumber(chrdata("ll_r"), 0) & "</TD>" & vbCrLf
                    Else
                        outtab1 += "<TD class='datacenter'>N/A</TD>" & vbCrLf
                    End If
                    If Not IsDBNull(chrdata("pi_r")) Then
                        outtab1 += "<TD class='datacenter'>" & FormatNumber(chrdata("pi_r"), 0) & "</TD>" & vbCrLf
                    Else
                        outtab1 += "<TD class='datacenter'>N/A</TD>" & vbCrLf
                    End If
                    outtab1 += "<TD class='datacenter'>&nbsp;</TD>" & vbCrLf
                Next
                'outtab1 += "</TABLE></TD></TR>" & vbCrLf

                'outtab1 += "<TR class='comptabrow2'><TD colspan='4' class='dataleft'>&nbsp;</TD></TR>" & vbCrLf
                'outtab1 += "<TR class='comptabrow2'><TD colspan='4'><TABLE bgcolor='lightcyan' width='100%' border='0' cellpadding='0' cellspacing='0' align='right'>" & vbCrLf
                outtab1 += "<TR class='comptabrow1'><TD colspan='1'>&nbsp;</td><TD colspan='14' class='dataleft'><B>Horizon Data: Physical Properties</B>* (representative values)</TD></TR>" & vbCrLf
                outtab1 += "<TR class='comptabrow1'>"
                outtab1 += "<TH colspan='1' width='1%' class='datacenter'>&nbsp;</TD>" & vbCrLf
                outtab1 += "<TH colspan='2' width='4%' class='ljustify'>Name </TD>" & vbCrLf
                outtab1 += "<TH colspan='2' width='8%' class='ljustify'>Depth</TD>" & vbCrLf
                outtab1 += "<TH colspan='2' width='27%' class='ljustify'>Texture</TD>" & vbCrLf
                outtab1 += "<TH colspan='2' width='26%' class='ljustify'>AASHTO Class</TD>" & vbCrLf
                outtab1 += "<TH colspan='2' width='26%' class='ljustify'>UNIFIED Class</TD>" & vbCrLf
                outtab1 += "<TH colspan='2' width='5%' class='ljustify'>Kf</TD>" & vbCrLf
                outtab1 += "<TH colspan='2' width='5%' align='left' valign='bottom' class='ljustify'><font face='arial' size='2'>Kw</TD>" & vbCrLf
                'outtab1 += "<TD colspan='1' width='2%' align='right' valign='bottom'><font face='arial' size='2'>&nbsp;</TD>" & vbCrLf

                For Each chrdata In chrdatatab.Rows
                    outtab1 += "<TR>"
                    outtab1 += "<TD colspan='1' width='1%' class='datacenter'>&nbsp;</TD>" & vbCrLf

                    outtab1 += "<TD class='ljustify' colspan='2'><B>" & chrdata("hzname") & "</B></TD>" & vbCrLf
                    outtab1 += "<TD class='ljustify' colspan='2'>" & vbCrLf
                    If Not IsDBNull(chrdata("hzdept_r")) Then
                        outtab1 += FormatNumber(chrdata("hzdept_r") / 2.54, 0)
                    End If
                    If Not IsDBNull(chrdata("hzdepb_r")) Then
                        outtab1 += " - " & FormatNumber(chrdata("hzdepb_r") / 2.54, 0) & vbCrLf
                    End If
                    outtab1 += " in</TD>" & vbCrLf

                    outtab1 += "<TD class='ljustify' colspan='2'>" & vbCrLf

                    Dim ctxSQL As String = "SELECT * FROM soils2019.chtexturegrp where chkey = '" & chrdata("chkey") & "';"
                    Dim ctxdatatable As DataTable = getSqlDataTable(ctxSQL)

                    If ctxdatatable.Rows.Count > 0 Then
                        Dim ct1 As Integer = 1
                        For Each ctxdata In ctxdatatable.Rows
                            If ct1 > 1 Then
                                outtab1 += ", " & vbCrLf
                            End If
                            If LCase(ctxdata("rvindicator")) = "yes" Then
                                outtab1 += "<B>" & ctxdata("texture") & "</B>" & vbCrLf
                            Else
                                outtab1 += ctxdata("texture") & vbCrLf
                            End If
                            If LCase(ctxdata("stratextsflag")) = "yes" Then
                                outtab1 += "(Stratified)" & vbCrLf
                            End If
                            ct1 = ct1 + 1
                        Next
                    Else
                        outtab1 += "N/A" & vbCrLf
                    End If
                    outtab1 += "</TD>" & vbCrLf

                    outtab1 += "<TD class='datacenter' colspan='2'>" & vbCrLf

                    Dim casSQL As String = "SELECT * FROM soils2019.chaashto where chkey = '" & chrdata("chkey") & "';"
                    Dim casdatatable As DataTable = getSqlDataTable(casSQL)
                    If casdatatable.Rows.Count > 0 Then
                        Dim ct1 As String = 1
                        For Each casdata In casdatatable.Rows
                            If ct1 > 1 Then
                                outtab1 += ", " & vbCrLf
                            End If
                            If LCase(casdata("rvindicator")) = "yes" Then
                                outtab1 += "<B>" & casdata("aashtocl") & "</B>" & vbCrLf
                            Else
                                outtab1 += casdata("aashtocl") & vbCrLf
                            End If
                            ct1 = ct1 + 1
                        Next
                    Else
                        outtab1 += "N/A" & vbCrLf
                    End If
                    outtab1 += "</TD>" & vbCrLf
                    outtab1 += "<TD class='datacenter' colspan='2'>" & vbCrLf

                    Dim cunSQL As String = "SELECT * FROM soils2019.chunified where chkey = '" & chrdata("chkey") & "';"
                    Dim cundatatable As DataTable = getSqlDataTable(cunSQL)
                    If cundatatable.Rows.Count > 0 Then
                        Dim ct1 As String = 1
                        For Each cundata In cundatatable.Rows
                            If ct1 > 1 Then
                                outtab1 += ", " & vbCrLf
                            End If
                            If LCase(cundata("rvindicator")) = "yes" Then
                                outtab1 += "<B>" & cundata("unifiedcl") & "</B>" & vbCrLf
                            Else
                                outtab1 += cundata("unifiedcl") & vbCrLf
                            End If
                            ct1 = ct1 + 1
                        Next
                    Else
                        outtab1 += "N/A" & vbCrLf
                    End If
                    outtab1 += "</TD>" & vbCrLf
                    If Not IsDBNull(chrdata("kwfact")) Then
                        outtab1 += "<TD class='datacenter' colspan='2'>" & chrdata("kffact") & "</TD>" & vbCrLf
                    Else
                        outtab1 += "<TD class='datacenter' colspan='2'>N/A</TD>" & vbCrLf
                    End If
                    If Not IsDBNull(chrdata("kffact")) Then
                        outtab1 += "<TD class='datacenter' colspan='2'>" & chrdata("kwfact") & "</TD>" & vbCrLf
                    Else
                        outtab1 += "<TD class='datacenter' colspan='2'>N/A</TD>" & vbCrLf
                    End If
                    'outtab1 += "<TD class='datacenter'>&nbsp;</TD></TR>" & vbCrLf
                Next
                'outtab1 += "<TR><TD colspan='14' class='ljustify'>* Bold type indicates the representative rating for the horizon.</TD>"
                'outtab1 += "<TD class='datacenter'>&nbsp;</TD></TR>" & vbCrLf
                'outtab1 += "</TABLE></TD></TR>" & vbCrLf
            End If

            outtab1 += "</table><br />"
            'report.Text += outtab2

            'Productivity information
            Dim ccrSQL As String = "SELECT * FROM soils2019.cocropyld where cokey = '" & cdata("cokey") & "';"
            Dim ccrdatatable As DataTable = getSqlDataTable(ccrSQL)

            outtab1 += "<div class='report_subheader'>Productivity Information</div>"
            outtab1 += "<table class='prodtab'>"

            Dim c As Integer = 1
            Dim Number2 As Integer = ccrdatatable.Rows.Count
            'outtab1 += "<TR bgcolor='lightcyan'><TD colspan='4' align='left' valign='center'><font face='arial' size='2'>&nbsp;</FONT></TD></TR>" & vbCrLf
            'outtab1 += "<TR class='prodtabrow1'><TD colspan='4'><TABLE class='prodtab'>" & vbCrLf
            outtab1 += "<TR class='prodtabrow1'><TD colspan='2' class='dataleft'><B>Estimated Crop Yields</B> (Acre)</TD>"
            outtab1 += "<TD colspan='2' class='dataleft'><B>Non-irrigated Land Capability Class:</B> " & cdata("nirrcapcl") & cdata("nirrcapscl") & "</TD></TR>" & vbCrLf
            If Not IsDBNull(cdata("foragesuitgrpid")) Then
                outtab1 += "<TR class='prodtabrow1'><TD colspan='2' class='dataleft'>&nbsp;</TD><TD colspan='2' class='dataleft'><B>Forage Suitability Group:</B> " & Replace(cdata("foragesuitgrpid"), "MO", "") & "</TD></TR>" & vbCrLf
            Else
                outtab1 += "<TR class='prodtabrow1'><TD colspan='2' class='dataleft'>&nbsp;</TD><TD colspan='2' class='dataleft'><B>Forage Suitability Group:</B> N/A</FONT></TD></TR>" & vbCrLf
            End If

            If ccrdatatable.Rows.Count > 0 Then
                For Each CCRdata As DataRow In ccrdatatable.Rows
                    If c Mod 2 = 1 Then
                        outtab1 += "<TR class='prodtabrow2'>" & vbCrLf
                    End If

                    'oops - no localplant table in the database any more?
                    'Dim plantSQL As String = "SELECT * from localplant where lplantname = '" & CCRdata("cropname") & "' ;"
                    'Dim plantdatatab As DataTable = getSqlDataTable(plantSQL)

                    'If plantdatatab.Rows.Count > 0 Then
                    'outtab1 += "<TD width='25%' colspan='1' class='dataleft'><B>" & CCRdata("cropname") & "</B> <font size='1'><a target='_blank' "
                    'outtab1 += "href='http://plants.usda.gov/java/profile?symbol=" & plantdatatab.Rows(0)("lplantsym") & "'>(" & plantdatatab.Rows(0)("lplantsciname") & ")</a></FONT></TD>" & vbCrLf
                    'Else
                    outtab1 += "<TD width='25%' class='dataleft'><B>" & CCRdata("cropname") & "</B></TD>" & vbCrLf
                    'End If
                    If Not IsDBNull(CCRdata("nonirryield_r")) Then
                        outtab1 += "<TD width='25%' class='dataleft'>" & FormatNumber(CCRdata("nonirryield_r"), 2) & " " & CCRdata("yldunits") & "</TD>" & vbCrLf
                    Else
                        outtab1 += "<TD width='25%' class='dataleft'>N/A</TD>" & vbCrLf
                    End If
                    If c Mod 2 = 0 Then
                        outtab1 += "</TR>" & vbCrLf
                    End If
                    c += 1
                Next
                If c Mod 2 = 0 Then
                    outtab1 += "<TD width='25%' class='dataleft'>&nbsp;</TD>" & vbCrLf
                    outtab1 += "<TD width='25%' class='dataleft'>&nbsp;</TD></TR>" & vbCrLf
                End If

                outtab1 += "<TR class='prodtabrow2'><TD colspan='4' class='dataleft'>&nbsp;</TD></TR>" & vbCrLf
                'outtab1 += "</TABLE></TD></TR>" & vbCrLf

                Dim cfpSQL As String = "SELECT * FROM soils2019.coforprod where cokey = '" & cdata("cokey") & "';"
                Dim cfpdatatab As DataTable = getSqlDataTable(cfpSQL)

                c = 1
                If cfpdatatab.Rows.Count > 0 Then
                    'outtab1 += "<TR class='prodtabrow1'><TD colspan='4'><TABLE class='prodtab'>" & vbCrLf
                    outtab1 += "<TR class='prodtabrow1'><TD colspan='2' class='dataleft'><B>Forest Productivity</B> - Site Index - Production (yr<sup>3</sup>/acre-year)</TD>"
                    outtab1 += "<TD colspan='2' class='dataleft'><B>Conservation Tree Shrub Group: </B>" & cdata("constreeshrubgrp") & "&nbsp;</TD></TR>" & vbCrLf
                    For Each cfpdata In cfpdatatab.Rows
                        If c Mod 2 = 1 Then
                            outtab1 += "<TR class='prodtabrow2'>" & vbCrLf
                        End If
                        outtab1 += "<TD width='31%' class='dataleft'><B>" & cfpdata("plantcomname") & "</B> <font size='1'><a target='_blank' "
                        outtab1 += "href='http://plants.usda.gov/java/profile?symbol=" & cfpdata("plantsym") & "'>(" & cfpdata("plantsciname") & ")</a></TD>" & vbCrLf
                        If IsNumeric(cfpdata("siteindex_r")) Then
                            If Not IsDBNull(cfpdata("fprod_r")) Then
                                outtab1 += "<TD width='19%' class='dataleft'>" & FormatNumber(cfpdata("siteindex_r"), 0) & " - " & FormatNumber(cfpdata("fprod_r"), 0) & "ft<sup>3</sup>/ac-yr</TD>" & vbCrLf
                            Else
                                outtab1 += "<TD width='19%' class='dataleft'>" & FormatNumber(cfpdata("siteindex_r"), 0) & "ft<sup>3</sup>/ac-yr</TD>" & vbCrLf
                            End If
                        Else
                            outtab1 += "<TD width='6%' class='dataleft'>Unrated</TD>" & vbCrLf
                        End If
                        If c Mod 2 = 0 Then
                            outtab1 += "</TR>" & vbCrLf
                        End If
                        c += 1
                    Next
                    If c Mod 2 = 0 Then
                        outtab1 += "<TD>&nbsp;</TD><TD>&nbsp;</TD></TR>" & vbCrLf
                    End If

                    If Not IsDBNull(cdata("constreeshrubgrp")) Then
                        Dim mtx2SQL As String = "SELECT * FROM soils2019.mutext where mukey = '" & mukey & "' and textcat = 'CTSG' and text like '% CTSG " & cdata("constreeshrubgrp") & "%' ;"
                        Dim mtx2datatab As DataTable = getSqlDataTable(mtx2SQL)
                        If mtx2datatab.Rows.Count = 1 Then
                            outtab1 += "<TR class='prodtabrow2'><TD colspan='4' class='dataleft'><BR>&nbsp;&nbsp;&nbsp; " & mtx2datatab.Rows(0)("text") & "</TD></TR>" & vbCrLf
                        End If
                    End If

                    outtab1 += "<TR class='prodtabrow2'><TD colspan='4' class='dataleft'>&nbsp;</TD></TR>" & vbCrLf
                    'outtab1 += "</TABLE></TD></TR>" & vbCrLf
                End If

                Dim cmgSQL As String = "SELECT * FROM soils2019.cotreestomng where cokey = '" & cdata("cokey") & "';"
                Dim cmgdatatab As DataTable = getSqlDataTable(cmgSQL)

                c = 1
                If cmgdatatab.Rows.Count > 0 Then
                    'outtab1 += "<TR class='prodtabrow2'><TD colspan='4' class='dataleft'>&nbsp;</TD></TR>" & vbCrLf
                    'outtab1 += "<TR class='prodtabrow1'><TD colspan='4'><TABLE class='prodtab'>" & vbCrLf
                    outtab1 += "<TR class='prodtabrow1'><TD colspan='2' class='dataleft'><B>Trees to Manage</B></TD></TR>" & vbCrLf
                    For Each cmgdata As DataRow In cmgdatatab.Rows
                        If c Mod 2 = 1 Then
                            outtab1 += "<TR class='prodtabrow2'>" & vbCrLf
                        End If
                        outtab1 += "<TD width='50%' class='dataleft'><B>" & cmgdata("plantcomname") & "</B> <font size='1'><a target='_blank' "
                        outtab1 += "href='http://plants.usda.gov/java/profile?symbol=" & cmgdata("plantsym") & "'>(" & cmgdata("plantsciname") & ") </a></TD>" & vbCrLf
                        If c Mod 2 = 0 Then
                            outtab1 += "</TR>" & vbCrLf
                        End If
                        c += 1
                    Next
                    If c Mod 2 = 0 Then
                        outtab1 += "<TD>&nbsp;</TD></TR>" & vbCrLf
                    End If
                    outtab1 += "<TR class='prodtabrow2'><TD colspan='4' class='dataleft'>&nbsp;</TD></TR>" & vbCrLf
                    'outtab1 += "</TABLE></TD></TR>" & vbCrLf
                End If

                Dim cwbSQL As String = "SELECT * FROM soils2019.copwindbreak where cokey = '" & cdata("cokey") & "';"
                Dim cwbdatatab As DataTable = getSqlDataTable(cwbSQL)

                c = 1
                If cwbdatatab.Rows.Count > 0 Then

                    'outtab1 += "<TR class='prodtabrow1'><TD colspan='4'><TABLE class='prodtab'>" & vbCrLf
                    outtab1 += "<TR class='prodtabrow1'><TD colspan='2' class='dataleft'><B>Windbreaks and Environmental Plantings</B> (representative heights at 20 yrs)</TD>"
                    If IsDBNull(cdata("wndbrksuitgrp")) Then
                        outtab1 += "<TD colspan='2' class='dataright'>&nbsp;</TD></TR>" & vbCrLf
                    Else
                        outtab1 += "<TD colspan='2' class='dataright'>Suitability Group: <B>" & cdata("wndbrksuitgrp") & "&nbsp;</B> </TD></TR>" & vbCrLf
                    End If
                    For Each cwbdata As DataRow In cwbdatatab.Rows
                        If c Mod 2 = 1 Then
                            outtab1 += "<TR class='prodtabrow2'>" & vbCrLf
                        End If
                        outtab1 += "<TD width='44%' class='dataleft'><B>" & cwbdata("plantcomname") & "</B> <font size='1'><a target='_blank' "
                        outtab1 += "href='http://plants.usda.gov/java/profile?symbol=" & cwbdata("plantsym") & "'>(" & cwbdata("plantsciname") & ")</a></TD>" & vbCrLf
                        If IsNumeric(cwbdata("wndbrkht_r")) Then
                            outtab1 += "<TD width='6%' class='dataright'>" & FormatNumber(cwbdata("wndbrkht_r") * 3.408, 0) & "ft&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>" & vbCrLf
                        Else
                            outtab1 += "<TD width='6%' class='dataleft'>Unrated</TD>" & vbCrLf
                        End If
                        If c Mod 2 = 0 Then
                            outtab1 += "</TR>" & vbCrLf
                        End If
                        c += 1
                    Next
                    If c Mod 2 = 0 Then
                        outtab1 += "<TD>&nbsp;</TD><TD>&nbsp;</TD></TR>" & vbCrLf
                    End If
                    outtab1 += "<TR class='prodtabrow2'><TD colspan='4' class='dataleft'>&nbsp;</TD></TR>" & vbCrLf
                    'outtab1 += "</TABLE></TD></TR>" & vbCrLf
                End If
            Else
                outtab1 += "<TR class='prodtabrow2'><TD colspan='4' class='dataleft'>Productivity information not available.</TD></TR>"
                'outtab1 += "</TABLE></TD></TR>" & vbCrLf
            End If

            outtab1 += "</table>"
            'report.Text += outtab1

            ' *** is the following still used? **************************************************************************************************************
            '       t2 = Split("wlgrain,wlgrass,wlherbaceous,wlshrub,wlconiferous,wlhardwood,wlwetplant,wlshallowwat", ",")
            '       b = 0
            '       For a = 0 To UBound(t2)
            '           If Not IsNull(cdata(t2(a))) Then
            '               b = b + 1
            '           End If
            '       Next
            '       If b > 0 Then
            '           t1 = Split("Grain Habitat,Grass Habitat,Herbaceous Habitat,Shrub Habitat,Conifer Habitat,Hardwood Habitat,Wetland Habitat,Water Habitat", ",")
            '           t3 = Split("Openland Wildlife,&nbsp;,Rangeland Wildlife,&nbsp;,Woodland Wildlife,&nbsp;,Wetland Wildlife,&nbsp;", ",")
            '           t4 = Split("wlopenland,&nbsp;,wlrangeland,&nbsp;,wlwoodland,&nbsp;,wlwetland,&nbsp;", ",")
            '           Response.Write("<TR bgcolor='wheat'><TD colspan='4' align='left' valign='center'><font face='arial' size='2'>&nbsp;</FONT></TD></TR>" & vbCrLf)
            '           Response.Write("<TR><TD colspan='4'><TABLE width='100%' border='0' cellpadding='0' cellspacing='0' bordercolor='darkgray' align='center' bgcolor='wheat'>" & vbCrLf)
            '           Response.Write("<TR bgcolor='tan'><TD colspan='2' align='left' valign='center'><font face='arial' size='2'><B>Wildlife and Habitat Management Suitability Ratings</B></FONT></TD>")
            '           '
            '           ' JLG 2/21/05 Needed an IsNull statement
            '           '
            '           If Not IsNUll(cdata("foragesuitgrpid")) Then
            '               Response.Write("<TD colspan='2' align='right' valign='center'><font face='arial' size='2'>Forage Suitability Group:  <B>" & Replace(cdata("foragesuitgrpid"), "MO", "") & "</B> </FONT></TD></TR>" & vbCrLf)
            '           Else
            '               Response.Write("<TD colspan='2' align='left' valign='center'><font face='arial' size='2'>Forage Suitability Group: <B>No Data Available</B> </FONT></TD></TR>" & vbCrLf)
            '           End If
            '           For a = 0 To UBound(t1)
            '               If Not IsNull(cdata(t2(a))) Then
            '                   Response.Write("<TR bgcolor='wheat'><TD width='25%' colspan='2' align='left' valign='center'><font face='arial' size='2'><B>" & t1(a) & "</B> - " & cdata(t2(a)) & "</FONT></TD>" & vbCrLf)
            '                   If InStr(t4(a), "nbsp") = 0 Then
            '                       If Not IsNull(cdata(t4(a))) Then
            '                           Response.Write("    <TD width='25%' colspan='2' align='left' valign='center'><font face='arial' size='2'><B>" & t3(a) & "</B>  - " & cdata(t4(a)) & "</FONT></TD></TR>" & vbCrLf)
            '                       Else
            '                           Response.Write("    <TD width='25%' colspan='2' align='left' valign='center'><font face='arial' size='2'>&nbsp;</FONT></TD>" & vbCrLf)
            '                       End If
            '                   Else
            '                       Response.Write("    <TD width='25%' colspan='2' align='left' valign='center'><font face='arial' size='2'>&nbsp;</FONT></TD>" & vbCrLf)
            '                   End If
            '               End If
            '           Next
            '           Response.Write("</TABLE></TD></TR>" & vbCrLf)
            '       End If

            '       CPLdata = Server.CreateObject("ADODB.Recordset")
            '       CPLdata.CursorLocation = 2
            '       cplSQL = "SELECT * FROM coeplants where cokey = '" & cdata("cokey") & "';"
            '       CPLdata.Open(cplSQL, conn, 1)
            '       If CPLdata.recordcount > 0 Then
            '           Number2 = CPLdata.recordcount
            '           CPLdata.MoveFirst()
            '           ct = 1
            '           Response.Write("<TR bgcolor='wheat'><TD colspan='4' align='left' valign='center'><font face='arial' size='2'>&nbsp;</FONT></TD></TR>" & vbCrLf)
            '           Response.Write("<TR><TD colspan='4'><TABLE width='100%' border='0' cellpadding='0' cellspacing='0' align='center'>" & vbCrLf)
            '           Response.Write("<TR bgcolor='tan'><TD align='left' valign='center'><font face='arial' size='2'><B>Common Plant Species</B> </FONT></TD>" & vbCrLf)
            '           If Not IsNUll(cdata("foragesuitgrpid")) Then
            '               Response.Write("<TD align='right' valign='center'><font face='arial' size='2'>Forage Suitability Group:  <B>" & Replace(cdata("foragesuitgrpid"), "MO", "") & "</B> </FONT></TD></TR>" & vbCrLf)
            '           Else
            '               Response.Write("<TD align='left' valign='center'><font face='arial' size='2'>Forage Suitability Group: <B>No Data Available</B> </FONT></TD></TR>" & vbCrLf)
            '           End If

            '           ct = 0
            '           Do While Not CPLdata.EOF
            '               ct = ct + 1
            '               If ct Mod 2 = 1 Then
            '                   Response.Write("<TR bgcolor='wheat'>" & vbCrLf)
            '               End If
            '               Response.Write("<TD width='50%' colspan='1' align='left' valign='center'><font face='arial' size='2'><B>" & CPLdata("plantcomname") & "</B><font size='1'><a target='_blank' href='http://plants.usda.gov/java/profile?symbol=" & CPLdata("plantsym") & "'>(" & CPLdata("plantsciname") & ")</a></font> </FONT></TD>" & vbCrLf)
            '               If ct Mod 2 = 0 Then
            '                   Response.Write("</TR>" & vbCrLf)
            '               End If
            '               CPLdata.MoveNext()
            '           Loop

            '           If ct Mod 2 = 1 Then
            '               Response.Write("<TD>&nbsp;</TD></TR>" & vbCrLf)
            '           End If
            '           If Not IsNull(cdata("foragesuitgrpid")) Then
            '               MTX3data = Server.CreateObject("ADODB.Recordset")
            '               MTX3data.CursorLocation = 2
            'mtx3SQL = "SELECT * FROM mutext where mukey = '"&mukey&"' and textcat = 'FSG' and text like '%Group "&replace(Cdata("foragesuitgrpid"),"MO","")&"%' ;"
            '               MTX3data.Open(mtx3SQL, conn, 3)
            '               If MTX3data.recordcount = 1 Then
            '                   Response.Write("<TR bgcolor='wheat'><TD colspan='2' align='left' valign='center'><font face='arial' size='2'>&nbsp;&nbsp;&nbsp; " & MTX3data("text") & "<br><br></TD></TR>" & vbCrLf)
            '               End If
            '           End If
            '           Response.Write("</TABLE></TD></TR>" & vbCrLf)

            'End If

            '*******************************************************************************************************************************************************

            '  BEGIN INTERPRETATION REPORT HERE

            'Dim outtab1 As String
            outtab1 += "<br />"
            outtab1 += "<div class='report_subheader'>Interpretations</div>"
            outtab1 += "<table class='noborders'>"

            Dim wplkey As String = mudata("lkey")
            Dim wpcokey As String = cdata("cokey")

            '  response.write "Load Time3: " & (timer - start_time) & " second(s), "&CINdata.recordcount&" "
            'outtab1 += "<TR class='inttabrow1'><TD colspan='4'><TABLE class='inttab'>" & vbCrLf
            outtab1 += "<TR class='inttabrow1' style='background-color:#c0c0c0;'><TD colspan='2' width='42%' class='dataleft'><B>Interpretations</B></TD>" & vbCrLf
            outtab1 += "<TD width='58%' colspan='2' class='dataleft'><B>Rating</B> (Score - Low - <B><i>Rep</i></b> - High )</TD></TR>" & vbCrLf
            'outtab1 += "<TR class='inttabrow1'><TD colspan='4'></TD></TR>" & vbCrLf
            c = 0
            Dim d As Integer = 0
            Dim bgcolor As String
            Dim lastrule As String = ""
            Dim rulename2 As String = ""
            Dim classtype1 As String = ""
            Dim classtype2 As String = ""
            Dim sfxcount As Integer = 0
            Dim classtitle As String = ""
            Dim intSQL As String
            Dim intdatatab As DataTable
            Dim classlist1 As Array = Split("AGR,AWM,BLM,DHS,ENG,FOR,GRL,MIL,NCCPI,URBREC,WMS,OTH", ",")
            Dim classlist2 As Array = Split("Agriculture,Animal Waste Management,Bureau of Land Management,Department of Homeland Security,Engineering,Forestry and Silviculture,Grazing and Range Management,Military,National Commodity Crop Productivity Index,Urban and Recreation,Water Management,Other Management Ratings", ",")

            For Each sfx In classlist1

                classtitle = classlist2(sfxcount)


                'Dim sqlStr As String = "Data Source=cliff;Initial Catalog=Soils2019;Persist Security Info=True;User ID=Soils2019;Password=9102slioS"
                'Dim conn As New SqlConnection(sqlStr)

                intSQL = "SELECT * from soils2019.COINTERP_" & Replace(sfx, "/", "") & " where cokey = '" & wpcokey & "' order by mrulename, seqnum;"


                'Dim sql = "INSERT INTO TESTLOG (SFX, SQLQUERY, THETIME) VALUES (@sfx, @sqlquery, @thetime)"
                'Dim comm As SqlCommand = Nothing
                'comm = New SqlCommand(sql, conn)
                'comm.Parameters.Add(New SqlParameter("@sfx", SqlDbType.Text))
                'comm.Parameters("@sfx").Value = sfx.ToString
                'comm.Parameters.Add(New SqlParameter("@sqlquery", SqlDbType.Text))
                'comm.Parameters("@sqlquery").Value = intSQL
                'comm.Parameters.Add(New SqlParameter("@thetime", SqlDbType.SmallDateTime))
                'comm.Parameters("@thetime").Value = Now
                'comm.Connection.Open()
                'comm.ExecuteNonQuery()
                'comm.Connection.Close()


                intdatatab = getSqlDataTable(intSQL)

                If intdatatab.Rows.Count > 0 Then
                    outtab1 += "<TR class='comptabrow1'><TD colspan='4' class='dataleft'><B>(" & sfx & ") - " & classtitle & "</B></td></TR>" & vbCrLf
                    'outtab1 += "<TR class='inttabrow1'><TD colspan='4'></TD></TR>" & vbCrLf
                End If

                For Each intdata As DataRow In intdatatab.Rows

                    Dim rname As Array = Split(intdata("mrulename"), "-")
                    If UBound(rname) = 2 Then
                        rname(1) = rname(1) + "-" + rname(2)
                    End If

                    If StrComp(Trim(intdata("mrulename")), Trim(intdata("rulename")), 0) = 0 Then
                        d = 0
                        c = c + 1
                        lastrule = intdata("mrulename")
                        If c Mod 2 = 1 Then
                            bgcolor = "inttabrow2" 'navahowhite
                        Else
                            bgcolor = "inttabrow3" 'blanchedalmond
                        End If
                        If UBound(rname) > 0 Then
                            outtab1 += "<TR style='border-top:solid 1px #dddddd;'><TD width='4%' colspan='2' class='ljustify'><B>" & c & "&nbsp;&nbsp;</B>" & vbCrLf
                            outtab1 += "<B>" & rname(1) & "</B></td>" & vbCrLf
                        Else
                            If sfx = "OTHER" Then
                                outtab1 += "<TR style='border-top:solid 1px #dddddd;'><TD width='4%' colspan='2' class='ljustify'><B>" & c & "&nbsp;&nbsp;</B>" & vbCrLf
                                outtab1 += "<B>" & intdata("mrulename") & "</B></td>" & vbCrLf
                            Else
                                outtab1 += "<TD width='42%' colspan='2' class='dataleft'><B>Database Error: No Rule Name in the database</B></td>" & vbCrLf
                            End If

                        End If

                        outtab1 += "<TD width='58%' colspan='2' class='dataleft'>" & vbCrLf

                        If IsDBNull(intdata("interplr")) Or IsDBNull(intdata("interphr")) Then
                            outtab1 += "Rating: " & intdata("interplrc")
                            If Not IsDBNull(intdata("interplr")) Then
                                outtab1 += " (<i><B>" & FormatNumber(intdata("interplr")) & "</B></i>)"
                            End If
                        Else
                            'If intdata("interplr") = intdata("interphr") Then
                            outtab1 += "Rating: " & intdata("interplrc")
                            If intdata("interpll") <> intdata("interphh") Then
                                outtab1 += " (" & FormatNumber(intdata("interpll")) & " - <i><B>" & FormatNumber(intdata("interplr")) & "</B></i> - " & FormatNumber(intdata("interphh")) & ")"
                            Else
                                outtab1 += " (<i><B>" & FormatNumber(intdata("interplr")) & "</B></i>)"
                            End If
                            ' Else
                            '    outtab1 += intdata("interplrc") & " (" & intdata("interplr") & " - " & intdata("interphrc") & " - " & intdata("interphr") & ")"
                            'End If
                        End If
                        outtab1 += "</TD></TR>" & vbCrLf
                    Else
                        If intdata("mrulename") = lastrule Then
                            d = d + 1
                            If UBound(rname) > 0 Or sfx = "OTHER" Then
                                outtab1 += "<TR><TD width='4%' class='dataleft' style='padding:0px;'>&nbsp;</td>" & vbCrLf
                                outtab1 += "<TD width='38%' class='dataleft' style='padding:0px;'>&nbsp;</td>" & vbCrLf
                            End If
                            outtab1 += "<TD width='58%' colspan='2' class='dataleftsmall' style='padding:0px;'>" & vbCrLf
                            If IsDBNull(intdata("interplr")) Or IsDBNull(intdata("interphr")) Then
                                outtab1 += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Reason " & d & ": " & intdata("interplrc")
                                If Not IsDBNull(intdata("interplr")) Then
                                    outtab1 += " (<i><B>" & FormatNumber(intdata("interplr")) & "</B></i>)"
                                End If
                            Else
                                If intdata("interplr") = intdata("interphr") Then
                                    outtab1 += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Reason " & d & ": " & intdata("interplrc")
                                    If intdata("interpll") <> intdata("interphh") Then
                                        outtab1 += " (" & FormatNumber(intdata("interpll")) & " - <i><B>" & FormatNumber(intdata("interplr")) & "</B></i> - " & FormatNumber(intdata("interphh")) & ")"
                                    Else
                                        outtab1 += " (<i><B>" & FormatNumber(intdata("interplr")) & "</B></i>)"
                                    End If

                                Else
                                    outtab1 += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Reason " & d & ": " & intdata("interplrc") & " (" & FormatNumber(intdata("interplr")) & ") - " & intdata("interphrc") & " (" & FormatNumber(intdata("interphr")) & ")"
                                End If
                            End If
                            outtab1 += "</TD></TR>" & vbCrLf
                        End If

                    End If

                    '              If c Mod 2 = 1 Then
                    '                  outtab1 += "</TR>" & vbCrLf
                    '              End If

                Next

                sfxcount += 1

            Next
            'outtab1 += "</TABLE></TD></TR>" & vbCrLf
            outtab1 += "</table></div>"
            'report.Text += outtab1

            '  END INTERPRETATION REPORT HERE
            'THIS is the end of the COMPONENT LOOP.  COMPONENT AND HORIZON REPORT MUST END BEFORE THIS END IF STATEMENT.
            'End If 'IF the COMPONENT DATA IS NOT NULL BASED ON HEL CLASSIFICATION

        Next



        Dim outtab5 As String = "</div></div><br /><br /><br /><hr /><div style='font-size:8pt;'>" & vbCrLf
        outtab5 += "The " & Ldata("mouagncyresp") & " is responsible for the <i>Soil Survey of " & Ldata("areaname") & "</i>. " & vbCrLf
        outtab5 += "The " & Ldata("mlraoffice") & ", MLRA Office is responsible for maintaining this legend. " & vbCrLf
        If IsDBNull(Ldata("areaacres")) Then
            outtab5 += "This 0 acre soil survey was " & vbCrLf
        Else
            outtab5 += "This " & FormatNumber(Ldata("areaacres"), 0) & " acre soil survey was " & vbCrLf
        End If
        If IsDBNull(Ldata("projectscale")) Then
            outtab5 += "conducted at an unknown scale." & vbCrLf
        Else
            outtab5 += "conducted at a scale of 1:" & FormatNumber(Ldata("projectscale"), 0) & "." & vbCrLf
        End If
        outtab5 += "This report was generated from a SQL Server database containing downloads from the NRCS Soils Data Mart " & vbCrLf
        outtab5 += "generated directly from the National Soil Information System (NASIS) by Cooperating State NRCS Offices. " & vbCrLf
        'outtab5 += "The <i>Soil Survey of " & Ldata("areaname") & "</i> was correlated " & LCase(Ldata("cordate")) & " and is" & vbCrLf
        'outtab5 += LCase(Ldata("legendcertstat")) & ".  The " & Ldata("mlraoffice") & ", MLRA Office is responsible for maintaining this legend,  " & vbCrLf
        'outtab5 += Ldata("mouagncyresp") & " is the agency responsible for this soil survey area.  This " & Left(Ldata("areaacres"), 3) & "," & Right(Ldata("areaacres"), 3) & " acre soil survey was " & vbCrLf
        'outtab5 += "conducted at a scale of 1:" & Left(Ldata("projectscale"), 2) & "," & Right(Ldata("projectscale"), 3) & ".  This report was generated from a SQL Server database containing downloads from the NRCS Soils Data mart, a Microsoft Access database generated " & vbCrLf
        'outtab5 += "directly from the National Soil Information System (NASIS) by Cooperating State NRCS Offices. " & vbCrLf
        'outtab5 += "CARES is a cooperator in the National Cooperative Soil Survey." & vbCrLf
        outtab5 += "<BR>Load Time: " & FormatNumber((Now - start_time).TotalSeconds.ToString, 4) & " seconds." & vbCrLf



        outtab5 += "<div style='margin-top:20px;'>Report provided by:<br /><br /><a href='http://cares.missouri.edu' target='_blank' title='Go to CARES website'><img src='../images/CARES_logo.png' width='300' /></a></div>"
        outtab5 += "</div>" & vbCrLf

        top.Text = outtab1
        footer.Text = outtab5
    End Sub

    Private Function getSqlDataTable(ByVal qStr As String) As DataTable
        Dim sqlStr As String = "Data Source=cliff;Initial Catalog=Soils2019;Persist Security Info=True;User ID=Soils2019;Password=9102slioS"
        Dim conn As New SqlConnection(sqlStr)
        Dim da As New SqlDataAdapter(qStr, conn)
        Dim dt As New DataTable
        da.Fill(dt)
        Return dt
    End Function
</script>



<!DOCTYPE html><html>
<head runat="server">
    <title>CARES | Soil Data Map Unit Interpretation Report</title>
    <style>
        body {
            font-family: "Open Sans", sans-serif;
        }
        table {
          border: 1px solid #ccc;
          border-collapse: collapse;
          margin: 0px 5px 0px 5px;
          padding: 0;
          width: 100%;
          table-layout: fixed;
        }

        table caption {
          font-size: 1.5em;
          margin: .5em 0 .75em;
        }

        table tr {
          background-color: #f0f0f0;
          border: 1px solid #ddd;
          padding: .35em;
        }

        table th,
        table td {
          padding: .625em;
          /*text-align: center;*/
        }

        table th {
          /*font-size: .85em;
          letter-spacing: .1em;
          text-transform: uppercase;*/
        }

        @media screen and (max-width: 600px) {
          table {
            border: 0;
          }

          table caption {
            font-size: 1.3em;
          }
  
          table thead {
            border: none;
            clip: rect(0 0 0 0);
            height: 1px;
            margin: -1px;
            overflow: hidden;
            padding: 0;
            position: absolute;
            width: 1px;
          }
  
          table tr {
            border-bottom: 3px solid #ddd;
            display: block;
            margin-bottom: .625em;
          }
  
          table td {
            border-bottom: 1px solid #ddd;
            display: block;
            font-size: .8em;
            text-align: right;
          }
  
          table td::before {
            /*
            * aria-label has no advantage, it won't be read inside a table
            content: attr(aria-label);
            */
            content: attr(data-label);
            float: left;
            font-weight: bold;
            text-transform: uppercase;
          }
  
          table td:last-child {
            border-bottom: 0;
          }
        }
        .h1
        {
           font-size: 20px;
           font-weight: bold;
           text-align: center;
           vertical-align: middle;           
        }
        .h2
        {
           font-size: 16px;
           font-weight: bold;
           text-align: left;
           vertical-align: middle;            
        }
        .h3
        {
           font-size: 14px;
           font-weight: bold;
           text-align: left;
           vertical-align: middle;            
        }
        .comptabrow1
        {
            background-color:#ddd;
        }
        .report_header
        {
            background-color:#ddd;
            font-weight:bold;
            font-size:14pt;
            padding:10px;
        }
        .report_content {
            padding:15px;
        }
        .inner_content {
            padding:10px;
        }
        .report_subheader {
            font-weight:bold;
            font-size:11pt;
            margin:5px 0px 5px 0px;
        }
        .accordion-toggle {
            cursor: pointer;
            background-color: #eee;
            color: #444;
            transition: 0.4s;
            font-weight:bold;
            padding:15px;
            margin-top:5px;
            font-size:12pt;
        }
        .accordion-content {
            display: none;
            padding:20px;
        }
        .accordion-content.default {
            display: block;
        }
        .accr_active, .accordion-toggle:hover {
            background-color: #ccc;
        }
        .accordion-toggle:after {
            content: '\002B';
            color: #777;
            font-weight: bold;
            float: right;
            margin-left: 5px;
        }
        .accr_active:after {
            content: "\2212";
        }
        .rjustify {
            text-align:right;
        }
        .ljustify {
            text-align:left;
        }
        .smallft {
            font-size:10pt;
        }
        .dataleft {
            text-align:left;
        }
        .noborders tr {
            border:0px;           
        }
        .dataleftsmall
        {
            font-size: 12px;
            font-weight: normal;
            text-align: left;
            vertical-align: middle;
            padding: 0px;
            margin: 0px;
        }

    </style>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script>
</head>
<body>

    <form id="form1" runat="server">
    <div>
        <asp:Literal ID="top" runat="server"></asp:Literal>
        <asp:Literal ID="report" runat="server"></asp:Literal>
        <asp:Literal ID="footer" runat="server"></asp:Literal>

        <script type="text/javascript">
            jQuery(document).ready(function ($) {
                $(".accr_header:first").addClass("accr_active");
                $('.accordion').find('.accordion-toggle').click(function () {
                    this.classList.toggle("accr_active");
                    //Expand or collapse this panel
                    $(this).next().slideToggle('fast');

                    //Hide the other panels
                    $(".accordion-content").not($(this).next()).slideUp('fast');
                    //console.log(this);
                    $(".accr_header").not($(this)).removeClass("accr_active");
                });
                var multcomps = $(".mult_comps").val();
                if (multcomps == 0) {
                    $(".accordion-content").css("display", "block");
                }
            });
        </script>
    </div>
    </form>
</body>
</html>
