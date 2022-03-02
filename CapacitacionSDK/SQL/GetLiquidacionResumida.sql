SELECT SS0."U_Code_Autor", SS0."U_Name_Autor",  SS0."FechaVencimiento",
SUM("Quantity") "Cantidad", SUM(SS0."ValorVendido") "ValorVendido", SUM(SS0."MontoRegalia") "MontoRegalia"
FROM (
	SELECT T4."U_Code_Autor", T4."U_Name_Autor", T0."DocCur", 
	CURRENT_DATE "FechaVencimiento", T1."Quantity", T1."LineTotal" "ValorVendido", T1."LineTotal" * T3."U_Porct_Cal" / 100 "MontoRegalia"
	FROM OINV T0
	INNER JOIN INV1 T1 ON T0."DocEntry" = T1."DocEntry"
	INNER JOIN OITM T2 ON T1."ItemCode" = T2."ItemCode" AND T2."U_HCO_Category" = '1'
	INNER JOIN "@LIQUIDACIONES_LN" T3 ON T1."ItemCode" = T3."U_Code_Libro"
	INNER JOIN "@LIQUIDACIONES_HD" T4 ON T3."DocEntry" = T4."DocEntry"
	WHERE T0."DocDate" BETWEEN '{0}' AND '{1}'
	AND T0."DocSubType" IN ('--', 'IB', 'DN')
	AND IFNULL("U_HCO_Liquidated", 'N') = 'N'
	UNION ALL
	SELECT T4."U_Code_Autor", T4."U_Name_Autor", T0."DocCur", 
	CURRENT_DATE "FechaVencimiento", -T1."Quantity", -T1."LineTotal" "ValorVendido", -1 * (T1."LineTotal" * T3."U_Porct_Cal" / 100) "MontoRegalia"
	FROM ORIN T0
	INNER JOIN RIN1 T1 ON T0."DocEntry" = T1."DocEntry"
	INNER JOIN OITM T2 ON T1."ItemCode" = T2."ItemCode" AND T2."U_HCO_Category" = '1'
	INNER JOIN "@LIQUIDACIONES_LN" T3 ON T1."ItemCode" = T3."U_Code_Libro"
	INNER JOIN "@LIQUIDACIONES_HD" T4 ON T3."DocEntry" = T4."DocEntry"
	WHERE T0."DocDate" BETWEEN '{0}' AND '{1}'
	AND IFNULL("U_HCO_Liquidated", 'N') = 'N'
) SS0
GROUP BY SS0."U_Code_Autor", SS0."U_Name_Autor", SS0."FechaVencimiento"