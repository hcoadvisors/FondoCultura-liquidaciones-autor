SELECT T4."U_Code_Autor" || ' - ' || T4."U_Name_Autor" "Autor", T0."DocEntry", T0."DocNum",
CASE T0."DocSubType"
	WHEN '--' THEN
		CASE T0."isIns"
			WHEN 'N' THEN 'Factura'
			WHEN 'Y' THEN 'Factura de reserva'
		END
	WHEN 'IB' THEN 'Boleta'
	WHEN 'DN' THEN 'Nota debito'
END "Tipo", T0."FolioPref", T0."FolioNum", T0."PTICode", T0."Letter", T0."FolNumFrom", T0."FolNumTo",
T1."ItemCode", T1."Dscription", T3."U_Porct_Obra", T0."DocDate", T0."DocCur", 
T1."Quantity", T1."LineTotal", T1."TotalFrgn", T3."U_Porct_Cal", T1."LineTotal" * T3."U_Porct_Cal" / 100 "MontoRegalia", T0."ObjType"
FROM OINV T0
INNER JOIN INV1 T1 ON T0."DocEntry" = T1."DocEntry"
INNER JOIN OITM T2 ON T1."ItemCode" = T2."ItemCode" AND T2."U_HCO_Category" = '1'
INNER JOIN "@LIQUIDACIONES_LN" T3 ON T1."ItemCode" = T3."U_Code_Libro"
INNER JOIN "@LIQUIDACIONES_HD" T4 ON T3."DocEntry" = T4."DocEntry"
WHERE T0."DocDate" BETWEEN '{0}' AND '{1}'
AND T0."DocSubType" IN ('--', 'IB', 'DN')
AND IFNULL("U_HCO_Liquidated", 'N') = 'N'
UNION ALL
SELECT T4."U_Code_Autor" || ' - ' || T4."U_Name_Autor" "Autor", T0."DocEntry", T0."DocNum", 'Nota credito' "Tipo",
T0."FolioPref", T0."FolioNum", T0."PTICode", T0."Letter", T0."FolNumFrom", T0."FolNumTo",
T1."ItemCode", T1."Dscription", T3."U_Porct_Obra", T0."DocDate", T0."DocCur", 
-T1."Quantity", - T1."LineTotal", - T1."TotalFrgn", - T3."U_Porct_Cal", -1 * (T1."LineTotal" * T3."U_Porct_Cal" / 100) "MontoRegalia", T0."ObjType"
FROM ORIN T0
INNER JOIN RIN1 T1 ON T0."DocEntry" = T1."DocEntry"
INNER JOIN OITM T2 ON T1."ItemCode" = T2."ItemCode" AND T2."U_HCO_Category" = '1'
INNER JOIN "@LIQUIDACIONES_LN" T3 ON T1."ItemCode" = T3."U_Code_Libro"
INNER JOIN "@LIQUIDACIONES_HD" T4 ON T3."DocEntry" = T4."DocEntry"
WHERE T0."DocDate" BETWEEN '{0}' AND '{1}'
AND IFNULL("U_HCO_Liquidated", 'N') = 'N'