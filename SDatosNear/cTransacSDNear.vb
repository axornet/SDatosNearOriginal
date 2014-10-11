Imports System.Data.SqlClient
Imports System.Web.Script.Serialization


Public Class cTransacSDNear

    Const IOS_CONDITION = " method like '%ios%' "
    Const PAYPAL_CONDITION = " (method like '%paypal%' or method like '%adaptivepayment%' or method like '%paypalrecurrent%'  or method like '%paypal_express%') and (method not like '%paypalexpresscheckoutprofilecreated%') "
    Const FREE_CONDITION = " method like '%free%' or method like '%contenttoolpayment%' "
    Const MANUAL_CONDITION = " method like '%purchaseorder%' or method like '%cashondelivery%'  or method like '%checkmo%' "
    Const AYCE_CONDITION = " method like '%aycepayment%' "

    Public Shared Function Generate( _
            ByRef pPgbGlobal As ProgressBar, _
            pPgbParcial As ProgressBar, _
            plblCurrentOp As Label, _
            plblTable As Label, _
            Optional ByRef pexError As Exception = Nothing) As String

        Dim sResulta As String = "Genear Transac SDNEAR" & vbCrLf
        Dim oCmd As SqlClient.SqlCommand

        sResulta += "Start all process" & Now.ToString & vbCrLf

        pPgbGlobal.Maximum = 100
        pPgbGlobal.Value = 0

        ProgressBarAdd(pPgbGlobal)
        oCmd = goConNear.CreateCommand
        oCmd.CommandTimeout = 999999
        plblCurrentOp.Text = "Deleting Records" : Application.DoEvents()
        sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
        oCmd.CommandText = "DELETE FROM TSD_Transactions"
        oCmd.ExecuteNonQuery()

        ProgressBarAdd(pPgbGlobal)
        oCmd = goConNear.CreateCommand
        oCmd.CommandTimeout = 999999
        plblCurrentOp.Text = "Clear TRC_code, payments, activepayment from User" : Application.DoEvents()
        sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
        oCmd.CommandText = "Update T_user set TRC_code = '',lastpaymentnumber=0 , activepaymentuser=0, lastunity = '' , lastregularity = 0, firstproductsale = '' "
        oCmd.ExecuteNonQuery()

        Dim lvstrColumns As String
        Dim sSpeacialSqlDelete As String
        Dim lvstrExpSql As String = <![CDATA[

SELECT 
0 as sd_source, 
sfoi.product_id as 'mage_id',
sfo.increment_id, sfo.grand_total, sfo.subtotal, sfo.discount_amount, sfo.total_refunded,sfo.coupon_code, cei.value as content_tool_customer_id, sfo.customer_email, 
sfo.customer_firstname, sfo.customer_lastname, sfo.created_at,  
sfoi.sku, sfoi.name, sfoi.price, 
(select rule_id from salesrule_coupon src where sfo.coupon_code = src.code limit 0,1) as 'coupon_rule_id',
(select sr.name from salesrule sr
 inner join salesrule_coupon src on src.rule_id = sr.rule_id
 where sfo.coupon_code = src.code limit 0,1) as 'coupon_rule_name',
case when freq.frequency is null 
  then sales_recurring_profile.period_frequency
  else freq.frequency  end as frequency,
case when  freq.period.period is null then sales_recurring_profile.period_unit
else freq.period.period end as period,
sfop.method,
case (SELECT cpei.value FROM catalog_product_entity_int cpei 
    WHERE sfoi.product_id = cpei.entity_id
    AND cpei.attribute_id = 147 LIMIT 1) 
WHEN 14 then 'License' 
WHEN 17 then 'License' 
WHEN 15 then 'Presentation'
WHEN 16 then 'Bundle'
WHEN 49 then 'AYCE'
WHEN 33 then 'Private Library'
end as 'product_type',
(SELECT cpei.value FROM catalog_product_entity_int cpei 
    WHERE sfoi.product_id = cpei.entity_id
    AND cpei.attribute_id = 144 LIMIT 1) as content_tool_product_id
FROM (select * from sales_flat_order sfo
 WHERE sfo.status = 'complete'  ) sfo
 INNER JOIN sales_flat_order_item sfoi 
  ON sfoi.order_id = sfo.entity_id  
 INNER JOIN customer_entity_int cei
  ON sfo.customer_id = cei.entity_id 
  AND cei.attribute_id = 141
 INNER JOIN sales_flat_order_payment sfop
  ON sfo.entity_id = sfop.parent_id
    LEFT JOIN
  (SELECT eaov.value as 'period', cpei.entity_id
   FROM catalog_product_entity_int cpei
   INNER JOIN eav_attribute_option_value eaov
   ON cpei.value = eaov.option_id
   WHERE attribute_id = 158) as period
 ON period.entity_id = sfoi.product_id
LEFT JOIN
  (SELECT eaov.value as 'frequency', cpei.entity_id
   FROM catalog_product_entity_int cpei
   INNER JOIN eav_attribute_option_value eaov
   ON cpei.value = eaov.option_id
   WHERE attribute_id = 159) as freq
 ON freq.entity_id = sfoi.product_id
LEFT JOIN sales_recurring_profile_order
ON     sales_recurring_profile_order.order_id = sfo.entity_id
LEFT JOIN sales_recurring_profile
  ON   sales_recurring_profile.profile_id = sales_recurring_profile_order.profile_id
order by increment_id desc

]]>.Value

        'SELECT 0 as sd_source, 
        'sfo.increment_id, sfo.grand_total, sfo.subtotal, sfo.discount_amount, sfo.total_refunded,sfo.coupon_code, cei.value as content_tool_customer_id, sfo.customer_email, sfo.customer_firstname, sfo.customer_lastname, sfo.created_at,  
        'sfoi.sku, sfoi.name, sfoi.price, 
        'case when freq.frequency is null 
        '  then sales_recurring_profile.period_frequency
        '  else freq.frequency  end as frequency,
        'case when  freq.period.period is null then sales_recurring_profile.period_unit
        'else freq.period.period end as period,
        'sfop.method,
        'case (SELECT cpei.value FROM catalog_product_entity_int cpei 
        '    WHERE sfoi.product_id = cpei.entity_id
        '    AND cpei.attribute_id = 147 LIMIT 1) 
        'WHEN 14 then 'License' 
        'WHEN 17 then 'License' 
        'WHEN 15 then 'Presentation'
        'WHEN 16 then 'Bundle'
        'end as 'product_type',
        '(SELECT cpei.value FROM catalog_product_entity_int cpei 
        '    WHERE sfoi.product_id = cpei.entity_id
        '    AND cpei.attribute_id = 144 LIMIT 1) as content_tool_product_id
        'FROM (select * from sales_flat_order sfo
        ' WHERE sfo.status = 'complete' 
        ' ) sfo
        ' INNER JOIN sales_flat_order_item sfoi 
        '  ON sfoi.order_id = sfo.entity_id  

        ' INNER JOIN customer_entity_int cei
        '  ON sfo.customer_id = cei.entity_id 
        '  AND cei.attribute_id = 141
        ' INNER JOIN sales_flat_order_payment sfop
        '  ON sfo.entity_id = sfop.parent_id
        '    LEFT JOIN
        '  (SELECT eaov.value as 'period', cpei.entity_id
        '   FROM catalog_product_entity_int cpei
        '   INNER JOIN eav_attribute_option_value eaov
        '   ON cpei.value = eaov.option_id
        '   WHERE attribute_id = 158) as period
        ' ON period.entity_id = sfoi.product_id
        'LEFT JOIN
        '  (SELECT eaov.value as 'frequency', cpei.entity_id
        '   FROM catalog_product_entity_int cpei
        '   INNER JOIN eav_attribute_option_value eaov
        '   ON cpei.value = eaov.option_id
        '   WHERE attribute_id = 159) as freq
        ' ON freq.entity_id = sfoi.product_id
        'LEFT JOIN sales_recurring_profile_order
        'ON     sales_recurring_profile_order.order_id = sfo.entity_id
        'LEFT JOIN sales_recurring_profile
        '  ON   sales_recurring_profile.profile_id = sales_recurring_profile_order.profile_id


        'Importa la base de datos
        ProgressBarAdd(pPgbGlobal)
        lvstrColumns = "sd_source, " & _
                       "mage_id, " & _
                       "increment_id, " & _
                       "grand_total, " & _
                       "subtotal, " & _
                       "discount_amount, " & _
                       "total_refunded, " & _
                       "coupon_code, " & _
                       "content_tool_customer_id, " & _
                       "customer_email, " & _
                       "customer_firstname, " & _
                       "sfo.customer_lastname, " & _
                       "sfo.created_at, " & _
                       "sku, " & _
                       "name, " & _
                       "price " & _
                       "method, " & _
                       "product_type, " & _
                       "content_tool_product_id, " & _
                       "presentations, " & _
                       "coupon_rule_id" & _
                       "coupon_rule_name"

        sSpeacialSqlDelete = "DELETE FROM T_Marketplace where sd_source = 0 or sd_source is NULL"
        'sResulta += gfstr_ImportaLinked(goConNear, pDataBase, lvstrExpSql, "T_Paypal", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)
        sResulta += gfstr_ImportaBulked(goConNear, goMagento, lvstrExpSql, "T_Marketplace", lvstrColumns, plblCurrentOp, plblTable, sSpeacialSqlDelete, pexError)


        sResulta = sResulta + Mkt_fact2000(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable)
        sResulta = sResulta + Mkt_fact2050(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable)
        sResulta = sResulta + Mkt_fact2100(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable)

        sResulta = sResulta + Mkt_fact3000(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable)
        sResulta = sResulta + Mkt_fact3050(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable)
        sResulta = sResulta + Mkt_5000(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable)
        sResulta = sResulta + Mkt_fact3100(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable)
        sResulta = sResulta + Mkt_fact3300(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable)

        sResulta = sResulta + Mkt_fact3150(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable)
        sResulta = sResulta + Mkt_fact3200(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable)
        sResulta = sResulta + Mkt_fact3250(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable)
        sResulta = sResulta + Mkt_fact3350(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable)

        'ayce presentaciones
        sResulta = sResulta + Mkt_fact3450(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable)

        'ayce 
        sResulta = sResulta + Mkt_fact2350(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable)

        'Private Library 
        sResulta = sResulta + Mkt_fact2400(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable)

        sResulta = sResulta + Mkt_BUNDLESVIAPRESENTATION(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable, IOS_CONDITION, "3050", "IOS")
        sResulta = sResulta + Mkt_BUNDLESVIAPRESENTATION(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable, PAYPAL_CONDITION, "3000", "PAYPAL")
        sResulta = sResulta + Mkt_BUNDLESVIAPRESENTATION(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable, FREE_CONDITION, "3100", "FREE")
        sResulta = sResulta + Mkt_BUNDLESVIAPRESENTATION(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable, MANUAL_CONDITION, "3300", "MANUAL")

        'ayce Bundels
        sResulta = sResulta + Mkt_BUNDLESVIAPRESENTATION(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable, AYCE_CONDITION, "3400", "MANUAL")


        sResulta = sResulta + GenerateUPHV2(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable, "0")

        'Refunded Viejo por que no se importa al viejo
        sResulta = sResulta + GeneratePaypalRefundedV2(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable, "0", "Nearpod Gold Edition", 2)

        'Importa UPH GOLD Referral Program
        sResulta = sResulta + GenerateUPHReferralV2(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable, "0")
        sResulta = sResulta + GenerateUPHReferralV2(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable, "1")


        'Calcula Cueota
        sResulta = sResulta + GeneratePaymentsAndWaivedV2(pPgbGlobal, pPgbParcial, plblCurrentOp, plblTable)

        'Esto esta sin resolver
        'ios	NULL	59  ---> no los estoy contando
        'paypal	NULL	9   ---> no los estoy contando
        'checkmo	License	9  --> no se lo que es
        'purchaseorder	License	7  ---> deberia ponerlo en factura MANUAL
        'cashondelivery	Presentation	3 ---> deberia ponerlo en factura MANUAL
        'waivedpayment	License	2 ---> NO SE QUE HACER TODAVIA
        'cashondelivery	License	1 ---> deberia ponerlo en factura MANUAL

        sResulta = sResulta + FixPresentationPrices(pPgbGlobal, pPgbParcial, plblCurrentOp)

        Call gp_InheritDate("TSD_Transactions", "created", "CR")

        sResulta += "End all process" & Now.ToString & vbCrLf

        Call GrabarLog(eLogType.eSFO, sResulta)

        Return sResulta


    End Function

    Public Shared Function FixPresentationPrices(ByRef pPgbGlobal As ProgressBar, _
                                                       pPgbParcial As ProgressBar, _
                                                       plblCurrentOp As Label)

        Dim oCmd As SqlClient.SqlCommand
        Dim sResulta As String

        ProgressBarAdd(pPgbGlobal)
        oCmd = goConNear.CreateCommand
        oCmd.CommandTimeout = 999999
        plblCurrentOp.Text = "Fix Presentation Prices From Transactions" : Application.DoEvents()
        sResulta = "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
        oCmd.CommandText = _
            "update T_presentation set price = TSD_Transactions.price " & _
            "   from T_Presentation " & _
            "   join T_presentation T2 on T2.id = T_Presentation.parentId and T_Presentation.FromStore = 1 " & _
            "   join TSD_Transactions on TSD_Transactions.userid = T_Presentation.userid and TSD_Transactions.presentationid = T_Presentation.parentid "
        oCmd.ExecuteNonQuery()
        Return sResulta


    End Function


    Public Shared Function GenerateUPHReferralV2(ByRef pPgbGlobal As ProgressBar, _
                                                       pPgbParcial As ProgressBar, _
                                                       plblCurrentOp As Label, _
                                                       plblTable As Label, _
                                                       psd_source As String, _
                                                       Optional ByRef pexError As Exception = Nothing) As String

        Dim oCmd As SqlClient.SqlCommand
        Dim sResulta As String = ""

        '***********************************************
        ' Primero todo Education 
        '
        '***********************************************
        Try

            '**********************************************************************************************
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Importing T_UserProductHistoric " : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "insert into Tsd_transactions (" & _
                                    "sd_source," & _
                                    "processRef, " & _
                                    "TRC_Code," & _
                                    "userId," & _
                                    "created," & _
                                    "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "units, " & _
                                    "source," & _
                                    "sourceId," & _
                                    "productId," & _
                                    "oldproductId," & _
                                    "presentationId," & _
                                    "regularity," & _
                                    "unity, " & _
                                    "upgradeAuthorizationManager, " & _
                                    "upgradeAuthorizationUser, " & _
                                    "upgradeAuthorizationMonths, " & _
                                    "expirationDate, " & _
                                    "upgradeAuthorizationUserId " & _
                            " ) " & _
                                "select " & _
                                    "sd_source, " & _
                                    "'GenerateUPHReferral' as processRef, " & _
                                    "'2150' as TRC_Code, " & _
                                    "userId, " & _
                                    "upgradetime as created, " & _
                                    "price," & _
                                    "price as grand_total," & _
                                    "price as subtotal," & _
                                    "0 as discount_amount," & _
                                    "0 as coupon_rule_id," & _
                                    "'' as coupon_rule_name," & _
                                    "'' as coupon_code," & _
                                    "1 as units, " & _
                                    "source, " & _
                                    "sourceid, " & _
                                    "productId, " & _
                                    "oldproductId," & _
                                    "0 as presentationId," & _
                                    "regularity, " & _
                                    "unity,  " & _
                                    "upgradeAuthorizationManager, " & _
                                    "upgradeAuthorizationUser, " & _
                                    "upgradeAuthorizationMonths, " & _
                                    "expirationDate, " & _
                                    "upgradeAuthorizationUserId " & _
                                    "from T_UserProductHistoric " & _
                                    "where source ='REFERRAL PROGRAM' "
            oCmd.ExecuteNonQuery()

        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf

        End Try

        Return sResulta

    End Function

    Public Shared Function GeneratePaypalRefundedV2(ByRef pPgbGlobal As ProgressBar, _
                                                          pPgbParcial As ProgressBar, _
                                                          plblCurrentOp As Label, _
                                                          plblTable As Label, _
                                                          psd_sorce As String,
                                                          pProduct As String,
                                                          pProductId As Integer,
                                                          Optional ByRef pexError As Exception = Nothing) As String


        Dim sResulta As String = ""
        Dim oCmd As SqlClient.SqlCommand

        Try

            '**********************************************************************************************
            'REFUND 
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Importing T_Paypal txt_type = '' AND " + pProduct + " and reason_code = refund or other and price Negative " : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "insert into Tsd_transactions (" & _
                                    "sd_source," & _
                                    "processRef, " & _
                                    "TRC_Code," & _
                                    "userId," & _
                                    "created," & _
                                    "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "units, " & _
                                    "source," & _
                                    "sourceId," & _
                                    "productId," & _
                                    "presentationId," & _
                                    "observ, " & _
                                    "regularity," & _
                                    "unity " & _
                                " ) " & _
                                "select " & _
                                    "sd_source, " & _
                                    "'GeneratePaypalqueNoesUPH 5' as processRef, " & _
                                    "'2200' as TRC_Code, " & _
                                    "userId, " & _
                                    "created, " & _
                                    "mc_gross as price, " & _
                                    "mc_gross as grand_total, " & _
                                    "mc_gross as subtotal, " & _
                                    "0 as discount_amount, " & _
                                    "0 as coupon_rule_id," & _
                                    "'' as coupon_rule_name," & _
                                    "'' as coupon_code," & _
                                    "-1 as units, " & _
                                    "'PAYPAL' as source, " & _
                                    "id as sourceid, " & _
                                    pProductId.ToString & " as productId, " & _
                                    "0 as presentationId," & _
                                    "payment_status as observ, " &
                                    "1 as regularity, " & _
                                    "'Y' as unity  " & _
                                    "from T_Paypal " & _
                                    "where  T_Paypal.txn_type = '' and payment_status in ('Refunded','Reversed') and " & _
                                    "       T_Paypal.productId = " + pProductId.ToString() + " and " & _
                                    "       T_Paypal.mc_gross < 0 and " & _
                                    "       sd_source = " + psd_sorce + " "
            oCmd.ExecuteNonQuery()



        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf

        End Try

        Return sResulta

    End Function

    Public Shared Function GenerateUPHV2(ByRef pPgbGlobal As ProgressBar, _
                                                       pPgbParcial As ProgressBar, _
                                                       plblCurrentOp As Label, _
                                                       plblTable As Label, _
                                                       psd_source As String, _
                                                       Optional ByRef pexError As Exception = Nothing) As String

        Dim oCmd As SqlClient.SqlCommand
        Dim sResulta As String = ""

        '***********************************************
        ' Primero todo Education 
        '
        '***********************************************
        Try

            '**********************************************************************************************
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Importing T_UserProductHistoric " : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "insert into Tsd_transactions (" & _
                                    "sd_source," & _
                                    "processRef, " & _
                                    "TRC_Code," & _
                                    "userId," & _
                                    "created," & _
                                    "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "units, " & _
                                    "source," & _
                                    "sourceId," & _
                                    "productId," & _
                                    "oldproductId," & _
                                    "presentationId," & _
                                    "regularity," & _
                                    "unity, " & _
                                    "upgradeAuthorizationManager, " & _
                                    "upgradeAuthorizationUser, " & _
                                    "upgradeAuthorizationMonths, " & _
                                    "expirationDate, " & _
                                    "upgradeAuthorizationUserId " & _
                            " ) " & _
                                "select " & _
                                    "sd_source, " & _
                                    "'GenerateUPH' as processRef, " & _
                                    "CASE " & _
                                    "  WHEN oldProductId = 3 and productId = 1 THEN '8000'  " & _
                                    "  WHEN oldProductId = 1000000004 and productId = 1000000001 THEN '8050'  " & _
                                    "  WHEN oldProductId = 3 and productId = 2 THEN '8300'  " & _
                                    "  WHEN oldProductId = 4 and productId = 3 THEN '8350'  " & _
                                    "  WHEN oldProductId = 2 and productId = 1 THEN '8100'  " & _
                                    "  WHEN oldProductId = 4 and productId = 1 THEN '8150'  " & _
                                    "  WHEN oldProductId = 1000000002 and productId = 1000000001 THEN '8200'  " & _
                                    "  WHEN oldProductId = 1000000003 and productId = 1000000001 THEN '8250'  " & _
                                    "  ELSE '9999'  " & _
                                    "End as TRC_Code, " & _
                                    "userId, " & _
                                    "upgradetime as created, " & _
                                    "price," & _
                                    "price as grand_total," & _
                                    "price as subtotal," & _
                                    "0 as discount_amount," & _
                                    "0 as coupon_rule_id," & _
                                    "'' as coupon_rule_name," & _
                                    "'' as coupon_code," & _
                                    "CASE " & _
                                    "  WHEN oldProductId = 3 and productId = 1 THEN -1  " & _
                                    "  WHEN oldProductId = 1000000004 and productId = 1000000001 THEN -1  " & _
                                    "  WHEN oldProductId = 3 and productId = 2 THEN -1   " & _
                                    "  WHEN oldProductId = 4 and productId = 3 THEN -1  " & _
                                    "  WHEN oldProductId = 2 and productId = 1 THEN -1  " & _
                                    "  WHEN oldProductId = 4 and productId = 1 THEN -1  " & _
                                    "  WHEN oldProductId = 1000000002 and productId = 1000000001 THEN -1  " & _
                                    "  WHEN oldProductId = 1000000003 and productId = 1000000001 THEN -1  " & _
                                    "  ELSE 0  " & _
                                    "End as units, " & _
                                    "source, " & _
                                    "sourceid, " & _
                                    "productId, " & _
                                    "oldproductId," & _
                                    "0 as presentationId," & _
                                    "regularity, " & _
                                    "unity,  " & _
                                    "upgradeAuthorizationManager, " & _
                                    "upgradeAuthorizationUser, " & _
                                    "upgradeAuthorizationMonths, " & _
                                    "expirationDate, " & _
                                    "upgradeAuthorizationUserId " & _
                                    "from T_UserProductHistoric " & _
                                    "where  sd_source = " + psd_source + " " & _
                                    "  and not( source = 'BACKEND' and upgradeAuthorizationManager = 'System' and upgradeAuthorizationUser = 'root') " & _
                                    "  and not ( " & _
                                    "       userid in (select userid from TSD_OriginalImport where waived = 'Y') and " & _
                                    "       (source = 'BACKEND' and upgradeAuthorizationManager = 'System') " & _
                                    "          )  " & _
                                    "  and not(source = 'PAYPAL' or source = 'APPLE' or source ='REFERRAL PROGRAM') " & _
                                    "  and (oldProductId > 0 and productId >0 and oldProductid <> productId)"

            oCmd.ExecuteNonQuery()

        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf

        End Try

        Return sResulta

    End Function


    Public Shared Function GeneratePaymentsAndWaivedV2(ByRef pPgbGlobal As ProgressBar, _
                                                      pPgbParcial As ProgressBar, _
                                                      plblCurrentOp As Label, _
                                                      plblTable As Label, _
                                                      Optional ByRef pexError As Exception = Nothing) As String

        Dim oCmd As SqlClient.SqlCommand
        Dim sResulta As String = ""

        ' 8125 Es para el downgrade Referral

        Dim cFACTURAS() = {"2000", "2050", "2100", "2150", "2200", "2250", "2300", "2350", "2400", "3000", "3050", "3100", "3150", "3200", "3250", "3300"}
        Dim cDOWNGRADES() = {"8000", "8050", "8100", "8125", "8150", "8200", "8250", "8300", "8350"}
        Dim cFACTURAS_PRODUCTO() = {"2000", "2050", "2100", "2150", "2200", "2250", "2300", "2350", "2400"}
        Dim cFACTURAS_CONTENIDO() = {"3000", "3050", "3100", "3150", "3200", "3250", "3300"}

        '***********************************************
        ' Primero todo Education 
        '
        '***********************************************
        Try

            '**********************************************************************************************
            'ProgressBarAdd(pPgbGlobal)
            'oCmd = goConNear.CreateCommand
            'oCmd.CommandTimeout = 999999
            'plblCurrentOp.Text = "Generating Payments  " : Application.DoEvents()
            'sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            'oCmd.CommandText = "select * from TSD_transactions "
            'Dim registros = oCmd.ExecuteReader()
            'For Each r In registros

            '    Console.WriteLine(r.userid)
            'Next

            'Dim a = 10

            Dim arraid As New ArrayList
            Dim arra2payments As New ArrayList
            Dim arra3userid As New ArrayList
            Dim arra4lastpaymentnumber As New ArrayList
            Dim arra5TRC_Code As New ArrayList
            Dim arra6activepaymentuser As New ArrayList
            Dim arra7unity As New ArrayList
            Dim arra8regularity As New ArrayList
            Dim arra9lastupgradeAuthorizationManager As New ArrayList
            Dim arra10lastupgradeAuthorizationUser As New ArrayList
            Dim arra11lastupgradeAuthorizationMonths As New ArrayList
            Dim arra12lastupgradeAuthorizationUserId As New ArrayList
            Dim arra13acumrevenueproduct As New ArrayList
            Dim arra14acumrevenuecontenido As New ArrayList

            Dim arra15id As New ArrayList
            Dim arra16erausuariopago As New ArrayList
            Dim arra17idChangeDowngradeToReferral As New ArrayList
            Dim arra18FirstProductSale As New ArrayList
            Dim arra19DateOfFirstUpgrade As New ArrayList

            Dim olduserid = 0
            Dim contpayments = 0
            Dim lastunity As String = ""
            Dim lastregularity = 0
            Dim firstProductSale = ""
            Dim dateOfFirstUpgrade As Date
            Dim lastTRC_Code As String = ""
            Dim lCommand As String = ""

            Dim lastupgradeAuthorizationManager As String = ""
            Dim lastupgradeAuthorizationUser As String = ""
            Dim lastupgradeAuthorizationMonths As Integer = 0
            Dim lastupgradeAuthorizationUserId As String = ""

            Dim acumrevenueproduct As Decimal = 0
            Dim acumrevenuecontenido As Decimal = 0

            Dim esUsuarioPago = "M"

            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "GeneratePaymentsAndWaivedV2" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf

            ' Prolijo un poco las cosas
            oCmd.CommandText = "Update TSD_Transactions " & _
                               "   set unity = 'Y' where UPPER(SUBSTRING(unity,1,1)) = 'Y' "
            oCmd.ExecuteNonQuery()
            oCmd.CommandText = "Update TSD_Transactions " & _
                               "   set unity = 'M' where UPPER(SUBSTRING(unity,1,1)) = 'M' "
            oCmd.ExecuteNonQuery()

            oCmd.CommandText = "Update TSD_Transactions " & _
                               "   set unity = 'Y' , regularity = 1 " & _
                               "   where unity = 'M' and regularity = 12 "
            oCmd.ExecuteNonQuery()

            'Dim Com As New SqlCommand("select * from TSD_Transactions where TRC_Code >='0300' and TRC_Code <='0349' order by userid asc, created asc  ", goConNear)
            Dim strSQl = "select * from TSD_Transactions " & _
                                      " where (TRC_Code in " & ArmoCadena(cFACTURAS) & " or TRC_Code in " & ArmoCadena(cDOWNGRADES) & ") " & _
                                      " order by userid asc, created asc  "
            Dim Com As New SqlCommand(strSQl, goConNear)

            Dim RDR = Com.ExecuteReader()
            If RDR.HasRows Then
                Do While RDR.Read

                    If (olduserid <> RDR.Item("userId")) Then

                        If olduserid <> 0 Then
                            arra3userid.Add(olduserid)
                            arra4lastpaymentnumber.Add(contpayments)
                            arra5TRC_Code.Add(lastTRC_Code)
                            arra7unity.Add(lastunity)
                            arra8regularity.Add(lastregularity)

                            arra18FirstProductSale.Add(firstProductSale)
                            arra19DateOfFirstUpgrade.Add(dateOfFirstUpgrade)

                            arra9lastupgradeAuthorizationManager.Add(lastupgradeAuthorizationManager)
                            arra10lastupgradeAuthorizationUser.Add(lastupgradeAuthorizationUser)
                            arra11lastupgradeAuthorizationMonths.Add(lastupgradeAuthorizationMonths)
                            arra12lastupgradeAuthorizationUserId.Add(lastupgradeAuthorizationUserId)

                            arra13acumrevenueproduct.Add(acumrevenueproduct)
                            arra14acumrevenuecontenido.Add(acumrevenuecontenido)

                            ' si la ultima fue una factura es un active payment user
                            If cFACTURAS_PRODUCTO.Contains(lastTRC_Code) Then
                                arra6activepaymentuser.Add(1)
                            Else
                                arra6activepaymentuser.Add(0)
                            End If

                        End If

                        esUsuarioPago = "N"
                        contpayments = 0
                        lastTRC_Code = ""
                        lastunity = ""
                        lastregularity = 0
                        firstProductSale = ""
                        dateOfFirstUpgrade = Nothing
                        acumrevenueproduct = 0
                        acumrevenuecontenido = 0
                        olduserid = RDR.Item("userId")
                    End If


                    If cFACTURAS_PRODUCTO.Contains(RDR.Item("TRC_Code")) Then
                        contpayments = contpayments + 1
                        arraid.Add(RDR.Item("id"))
                        arra2payments.Add(contpayments)
                        If Not (IsDBNull(RDR.Item("unity"))) Then
                            If RDR.Item("unity") = "M" Then
                                esUsuarioPago = "M"
                            Else
                                esUsuarioPago = "Y"
                            End If
                        End If
                    End If

                    If cFACTURAS_PRODUCTO.Contains(RDR.Item("TRC_Code")) Then
                        acumrevenueproduct = acumrevenueproduct + IIf(IsDBNull(RDR.Item("price")), 0, RDR.Item("price"))
                    End If

                    If cFACTURAS_CONTENIDO.Contains(RDR.Item("TRC_Code")) Then
                        acumrevenuecontenido = acumrevenuecontenido + IIf(IsDBNull(RDR.Item("price")), 0, RDR.Item("price"))
                    End If


                    If cFACTURAS_PRODUCTO.Contains(RDR.Item("TRC_Code")) Then

                        lastTRC_Code = RDR.Item("TRC_Code")
                        lastregularity = IIf(IsDBNull(RDR.Item("regularity")), 0, RDR.Item("regularity"))
                        lastunity = IIf(IsDBNull(RDR.Item("unity")), "", RDR.Item("unity"))

                        ' Guardo la primera factura de producto que se le hizo
                        If firstProductSale = "" Then
                            firstProductSale = RDR.Item("TRC_Code")
                            dateOfFirstUpgrade = RDR.Item("created")
                        End If

                    End If

                    If (cDOWNGRADES.Contains(RDR.Item("TRC_Code")) And esUsuarioPago <> "N") Then
                        arra15id.Add(RDR.Item("id"))
                        arra16erausuariopago.Add(esUsuarioPago)
                    End If

                    'referral downgrade
                    If (cDOWNGRADES.Contains(RDR.Item("TRC_Code")) And lastTRC_Code = "2150") Then
                        arra17idChangeDowngradeToReferral.Add(RDR.Item("id"))
                    End If


                Loop
            End If
            RDR.Close()

            plblCurrentOp.Text = "update TSD_Transactions set paymentnumber" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf


            lCommand = ""
            For i = 0 To arraid.Count - 1
                lCommand = lCommand & "update TSD_Transactions set paymentnumber = " & arra2payments.Item(i).ToString() & " where id = " & arraid.Item(i).ToString() & "; "

                If (i / 20 = Int(i / 20) Or i = arraid.Count - 1) Then
                    oCmd = goConNear.CreateCommand
                    oCmd.CommandTimeout = 999999
                    oCmd.CommandText = lCommand
                    oCmd.ExecuteNonQuery()
                    lCommand = ""
                End If

            Next

            plblCurrentOp.Text = "update TSD_Transactions set unity" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf

            lCommand = ""
            For i = 0 To arra15id.Count - 1
                lCommand = lCommand & "update TSD_Transactions set unity = '" & arra16erausuariopago.Item(i).ToString() & "', paidUser = '" & arra16erausuariopago.Item(i).ToString() & "' where id = " & arra15id.Item(i).ToString() & "; "

                If (i / 20 = Int(i / 20) Or i = arra15id.Count - 1) Then
                    oCmd = goConNear.CreateCommand
                    oCmd.CommandTimeout = 999999
                    oCmd.CommandText = lCommand
                    oCmd.ExecuteNonQuery()
                    lCommand = ""
                End If
            Next


            ' Cambia los referral downgrades al codigo correcto 
            lCommand = ""
            For i = 0 To arra17idChangeDowngradeToReferral.Count - 1
                lCommand = lCommand & "update TSD_Transactions set trc_code = '8125' where id = " & arra17idChangeDowngradeToReferral.Item(i).ToString() & "; "

                If (i / 20 = Int(i / 20) Or i = arra17idChangeDowngradeToReferral.Count - 1) Then
                    oCmd = goConNear.CreateCommand
                    oCmd.CommandTimeout = 999999
                    oCmd.CommandText = lCommand
                    oCmd.ExecuteNonQuery()
                    lCommand = ""
                End If

            Next


            plblCurrentOp.Text = "update T_User" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf

            For i = 0 To arra3userid.Count - 1

                Dim d1 As Date = arra19DateOfFirstUpgrade.Item(i)
                Dim limitedate As Date

                limitedate = New DateTime(2000, 1, 1)
                If d1 < limitedate Then
                    d1 = limitedate
                End If

                lCommand = lCommand & "update T_User " & _
                                   "    set TRC_Code = '" & arra5TRC_Code.Item(i) & "', " & _
                                   "        lastpaymentnumber = " & arra4lastpaymentnumber.Item(i).ToString & ", " & _
                                   "        activepaymentuser = " & arra6activepaymentuser.Item(i).ToString & ", " & _
                                   "        lastunity = '" & arra7unity.Item(i).ToString & "', " & _
                                   "        lastregularity = " & arra8regularity.Item(i).ToString & ", " & _
                                   "        firstproductsale = '" & arra18FirstProductSale.Item(i).ToString & "', " & _
                                   "        sdn_Stage30Date = '" & arra19DateOfFirstUpgrade.Item(i) & "', " & _
                                   "        sdn_Stage30 = case when cast('" & d1.ToString("yyyy-MM-dd HH:MM:ss") & "'  as datetime) > cast('2000-01-01' as datetime) then 1 else 0 end, " & _
                                   "        lastupgradeAuthorizationManager = '" & arra9lastupgradeAuthorizationManager.Item(i).ToString.Replace("'", String.Empty) & "', " & _
                                   "        lastupgradeAuthorizationUser = '" & arra10lastupgradeAuthorizationUser.Item(i).ToString & "', " & _
                                   "        lastupgradeAuthorizationMonths = " & arra11lastupgradeAuthorizationMonths.Item(i).ToString & ", " & _
                                   "        lastupgradeAuthorizationUserId = '" & arra12lastupgradeAuthorizationUserId.Item(i).ToString & "', " & _
                                   "        acumrevenueproduct = '" & arra13acumrevenueproduct.Item(i).ToString & "', " & _
                                   "        acumrevenuecontenido = '" & arra14acumrevenuecontenido.Item(i).ToString & "' " & _
                                   " where id = " & arra3userid.Item(i).ToString() & "; "

                If (i / 20 = Int(i / 20) Or i = arraid.Count - 1) Then
                    oCmd = goConNear.CreateCommand
                    oCmd.CommandTimeout = 999999
                    oCmd.CommandText = lCommand
                    oCmd.ExecuteNonQuery()
                    lCommand = ""
                End If
            Next

            plblCurrentOp.Text = "End Updates" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf

        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf

        End Try

        Return sResulta

    End Function

    Public Shared Function Mkt_BUNDLESVIAPRESENTATION(ByRef pPgbGlobal As ProgressBar, _
                                             pPgbParcial As ProgressBar, _
                                             plblCurrentOp As Label, _
                                             plblTable As Label,
                                             pCondition As String,
                                             pTRCCode As String,
                                             pSource As String,
                                             Optional ByRef pexError As Exception = Nothing) As String

        Dim oCmd As SqlClient.SqlCommand
        Dim sResulta As String = ""
        Dim lCommand As String = ""
        Dim arrCommand As New ArrayList


        Dim debugDato As String

        ProgressBarAdd(pPgbGlobal)
        oCmd = goConNear.CreateCommand
        oCmd.CommandTimeout = 999999
        plblCurrentOp.Text = "Explode Bundles" : Application.DoEvents()
        sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf

        Try

            'Dim strSQl = "select * from T_Marketplace " & _
            '             "where  T_Marketplace.sd_source = 0  " & _
            '             "       and ( " & pCondition & " )  " & _
            '             "       and (product_type = 'Bundle') "


            Dim strSQl = <![CDATA[
SELECT M.*,
   presentations2=STUFF((SELECT ','+ CONVERT(varchar, presentationID)  FROM T_BundlePresentation WHERE T_BundlePresentation.bundleId=M.mage_id FOR XML PATH('')) , 1 , 1 , '' )
FROM 
   T_MarketPlace M
Where M.product_type = 'Bundle'
and M.sd_source = 0 
and ( ]]>.Value & pCondition & <![CDATA[ ) ]]>.Value


            Dim Com As New SqlCommand(strSQl, goConNear)
            Dim RDR = Com.ExecuteReader()
            If RDR.HasRows Then
                Do While RDR.Read
                    Dim phrase As String
                    If IsDBNull(RDR.Item("presentations2")) Then
                        phrase = ""
                    Else
                        phrase = IIf(IsDBNull(RDR.Item("presentations2")), "", RDR.Item("presentations2"))
                    End If

                    Dim presentations() As String
                    presentations = phrase.Split({","}, StringSplitOptions.RemoveEmptyEntries)

                    If True Then
                        If (presentations.Count > 0) Then
                            Dim priceToSave = RDR.Item("Price") / presentations.Count

                            For i = 0 To presentations.Count - 1
                                debugDato = RDR.Item("increment_id")

                                Dim nameString As String = RDR.Item("name")

                                nameString = nameString.Replace("'", "")

                                lCommand = "insert into Tsd_transactions (" & _
                                                                "sd_source," & _
                                                                "processRef, " & _
                                                                "TRC_Code," & _
                                                                "userId," & _
                                                                "created," & _
                                                                "price," & _
                                                                "grand_total," & _
                                                                "subtotal," & _
                                                                "discount_amount," & _
                                                                "coupon_rule_id," & _
                                                                "coupon_rule_name," & _
                                                                "coupon_code," & _
                                                                "units, " & _
                                                                "source," & _
                                                                "sourceId," & _
                                                                "productId," & _
                                                                "presentationId, " & _
                                                                "regularity," & _
                                                                "unity, " & _
                                                                "observ, " & _
                                                                "bundleId " & _
                                                        " ) VALUES (" & _
                                                            "0, " & _
                                                            "'Explode Bundles', " & _
                                                            pTRCCode & ", " & _
                                                            RDR.Item("content_tool_customer_id") & ", " & _
                                                            "'" & RDR.Item("created_at") & "', " & _
                                                            priceToSave & " , " & _
                                                            priceToSave & " , " & _
                                                            priceToSave & " , " & _
                                                            "0, " & _
                                                            IIf(IsDBNull(RDR.Item("coupon_rule_id")), 0, RDR.Item("coupon_rule_id")) & ", " & _
                                                            "'" & IIf(IsDBNull(RDR.Item("coupon_rule_name")), "", RDR.Item("coupon_rule_name")) & "', " & _
                                                            "'" & IIf(IsDBNull(RDR.Item("coupon_code")), "", RDR.Item("coupon_code")) & "', " & _
                                                            "1, " & _
                                                            "'" & pSource & "', " & _
                                                            RDR.Item("increment_id") & ", " & _
                                                            "0, " & _
                                                            presentations(i) & ", " & _
                                                            IIf(IsDBNull(RDR.Item("frequency")), 0, RDR.Item("frequency")) & ", " & _
                                                            IIf(IsDBNull(RDR.Item("period")), 0, RDR.Item("period")) & ", " & _
                                                            "'" & nameString & "', " & _
                                                            RDR.Item("mage_id") & " " & _
                                                        " ) "


                                arrCommand.Add(lCommand)

                            Next

                        End If
                    End If

                Loop
            End If
            RDR.Close()

            For i = 0 To arrCommand.Count - 1

                oCmd = goConNear.CreateCommand
                oCmd.CommandTimeout = 999999
                oCmd.CommandText = arrCommand(i)
                oCmd.ExecuteNonQuery()
                lCommand = ""

            Next


            'plblCurrentOp.Text = "update TSD_Transactions set paymentnumber" : Application.DoEvents()
            'sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf

            'lCommand = ""
            'For i = 0 To arraid.Count - 1
            '    lCommand = lCommand & "update TSD_Transactions set paymentnumber = " & arra2payments.Item(i).ToString() & " where id = " & arraid.Item(i).ToString() & "; "

            '    If (i / 20 = Int(i / 20) Or i = arraid.Count - 1) Then
            '        oCmd = goConNear.CreateCommand
            '        oCmd.CommandTimeout = 999999
            '        oCmd.CommandText = lCommand
            '        oCmd.ExecuteNonQuery()
            '        lCommand = ""
            '    End If

            'Next


            plblCurrentOp.Text = "End Explode Bundles" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf

        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf

        End Try

        Return sResulta



    End Function

    Public Shared Function ProcessJson(xstr As String) As ArrayList


        Dim returnArray As New ArrayList

        Dim cadena As String
        Dim p, p2, p3, p4, p5 As Integer

        p = InStr(xstr, "XX01")
        Do While (p > 0)
            cadena = Mid(xstr, p, 20)
            p2 = InStr(p + 7, xstr, Chr(34))
            If p2 > 0 Then
                returnArray.Add(Mid(xstr, p + 7, p2 - (p + 7)))
            End If

            p = InStr(p + 7, xstr, "XX01")
        Loop

        Return returnArray

        'Dim ds = New JavaScriptSerializer()
        'ds.MaxJsonLength = 6097152
        'Dim j As Object = ds.Deserialize(Of Object)(json)

        'Try

        '    If (j.Count >= 2) Then
        '        Dim aLimit = j("presentations").Length
        '        If (j("presentations").Length > 0) Then
        '            For i = 0 To aLimit - 1
        '                returnArray.Add(j("presentations")(i)("external_id"))
        '            Next
        '        End If
        '    End If

        'Catch ex As Exception
        '    Return (New ArrayList)
        'End Try

        'Return returnArray


    End Function


    Public Shared Function Mkt_fact3250(ByRef pPgbGlobal As ProgressBar, _
                                             pPgbParcial As ProgressBar, _
                                             plblCurrentOp As Label, _
                                             plblTable As Label,
                                             Optional ByRef pexError As Exception = Nothing) As String

        Dim oCmd As SqlClient.SqlCommand
        Dim sResulta As String = ""

        Try


            '**********************************************************************************************
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Importing 3250 Bundles Free" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "insert into Tsd_transactions (" & _
                                    "sd_source," & _
                                    "processRef, " & _
                                    "TRC_Code," & _
                                    "userId," & _
                                    "created," & _
                                    "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "units, " & _
                                    "source," & _
                                    "sourceId," & _
                                    "productId," & _
                                    "presentationId, " & _
                                    "regularity," & _
                                    "unity, " & _
                                    "observ, " & _
                                    "bundleId " & _
                                " ) " & _
                                "select " & _
                                    "T_Marketplace.sd_source, " & _
                                    "'3250 bundle Free', " & _
                                    "'3250' as TRC_Code, " & _
                                    "content_tool_customer_id, " & _
                                    "created_at, " & _
                                    "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "1 as units, " & _
                                    "'FREE' as source, " & _
                                    "increment_id as sourceid, " & _
                                    "0, " & _
                                    "0, " & _
                                    "frequency, " & _
                                    "period,  " & _
                                    "name as observ, " & _
                                    "T_Marketplace.mage_id as bundleId " & _
                                    "from T_Marketplace " & _
                                    "where  T_Marketplace.sd_source = 0  " & _
                                    "       and ( " & FREE_CONDITION & " )  " & _
                                    "       and (product_type = 'Bundle') "
            'oscar octubre/2014 modifique lo de mageid 
            oCmd.ExecuteNonQuery()


        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf

        End Try

        Return sResulta

    End Function

    Public Shared Function Mkt_fact3150(ByRef pPgbGlobal As ProgressBar, _
                                             pPgbParcial As ProgressBar, _
                                             plblCurrentOp As Label, _
                                             plblTable As Label,
                                             Optional ByRef pexError As Exception = Nothing) As String

        Dim oCmd As SqlClient.SqlCommand
        Dim sResulta As String = ""

        Try


            '**********************************************************************************************
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Importing 3150 Bundles PAYPAL" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "insert into Tsd_transactions (" & _
                                    "sd_source," & _
                                    "processRef, " & _
                                    "TRC_Code," & _
                                    "userId," & _
                                    "created," & _
                                    "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "units, " & _
                                    "source," & _
                                    "sourceId," & _
                                    "productId," & _
                                    "presentationId, " & _
                                    "regularity," & _
                                    "unity, " & _
                                    "observ, " & _
                                    "bundleId " & _
                                " ) " & _
                                "select " & _
                                    "T_Marketplace.sd_source, " & _
                                    "'3150 bundle PAYPAL', " & _
                                    "'3150' as TRC_Code, " & _
                                    "content_tool_customer_id, " & _
                                    "created_at, " & _
                                    "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "1 as units, " & _
                                    "'PAYPAL' as source, " & _
                                    "increment_id as sourceid, " & _
                                    "0, " & _
                                    "0, " & _
                                    "frequency, " & _
                                    "period,  " & _
                                    "name as observ, " & _
                                    "mage_id as bundleId " & _
                                    "from T_Marketplace " & _
                                    "where  T_Marketplace.sd_source = 0  " & _
                                    "       and ( " & PAYPAL_CONDITION & " )  " & _
                                    "       and (product_type = 'Bundle') "

            'oscar Octubre/2014 Arriba modifique mage_id
            oCmd.ExecuteNonQuery()


        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf

        End Try

        Return sResulta

    End Function

    Public Shared Function Mkt_fact3350(ByRef pPgbGlobal As ProgressBar, _
                                            pPgbParcial As ProgressBar, _
                                            plblCurrentOp As Label, _
                                            plblTable As Label,
                                            Optional ByRef pexError As Exception = Nothing) As String

        Dim oCmd As SqlClient.SqlCommand
        Dim sResulta As String = ""

        Try


            '**********************************************************************************************
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Importing 3350 Bundles MANUAL" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "insert into Tsd_transactions (" & _
                                    "sd_source," & _
                                    "processRef, " & _
                                    "TRC_Code," & _
                                    "userId," & _
                                    "created," & _
                                    "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "units, " & _
                                    "source," & _
                                    "sourceId," & _
                                    "productId," & _
                                    "presentationId, " & _
                                    "regularity," & _
                                    "unity, " & _
                                    "observ, " & _
                                    "bundleId " & _
                                " ) " & _
                               "select " & _
                                    "T_Marketplace.sd_source, " & _
                                    "'3350 bundle MANUAL', " & _
                                    "'3350' as TRC_Code, " & _
                                    "content_tool_customer_id, " & _
                                    "created_at, " & _
                                    "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "1 as units, " & _
                                    "'MANUAL' as source, " & _
                                    "increment_id as sourceid, " & _
                                    "0, " & _
                                    "0, " & _
                                    "frequency, " & _
                                    "period,  " & _
                                    "name as observ, " & _
                                    "mage_id as bundleId " & _
                                    "from T_Marketplace " & _
                                    "where  T_Marketplace.sd_source = 0  " & _
                                    "       and ( " & MANUAL_CONDITION & " )  " & _
                                    "       and (product_type = 'Bundle') "
            'oscar Octubre/2014 Arriba modifique mage_id
            oCmd.ExecuteNonQuery()


        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf

        End Try

        Return sResulta

    End Function

    Public Shared Function Mkt_fact3200(ByRef pPgbGlobal As ProgressBar, _
                                            pPgbParcial As ProgressBar, _
                                            plblCurrentOp As Label, _
                                            plblTable As Label,
                                            Optional ByRef pexError As Exception = Nothing) As String

        Dim oCmd As SqlClient.SqlCommand
        Dim sResulta As String = ""

        Try


            '**********************************************************************************************
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Importing 3200 Bundles IOS" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "insert into Tsd_transactions (" & _
                                    "sd_source," & _
                                    "processRef, " & _
                                    "TRC_Code," & _
                                    "userId," & _
                                    "created," & _
                                    "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "units, " & _
                                    "source," & _
                                    "sourceId," & _
                                    "productId," & _
                                    "presentationId, " & _
                                    "regularity," & _
                                    "unity, " & _
                                    "observ, " & _
                                    "bundleId " & _
                                " ) " & _
                                 "select " & _
                                    "T_Marketplace.sd_source, " & _
                                    "'3200 bundle IOS', " & _
                                    "'3200' as TRC_Code, " & _
                                    "content_tool_customer_id, " & _
                                    "created_at, " & _
                                    "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "1 as units, " & _
                                    "'IOS' as source, " & _
                                    "increment_id as sourceid, " & _
                                    "0, " & _
                                    "0, " & _
                                    "frequency, " & _
                                    "period,  " & _
                                    "name as observ, " & _
                                    "mage_id as bundleId " & _
                                    "from T_Marketplace " & _
                                    "where  T_Marketplace.sd_source = 0  " & _
                                    "       and ( " & IOS_CONDITION & " )  " & _
                                    "       and (product_type = 'Bundle') "
            'oscar Octubre/2014 Arriba modifique mage_id
            oCmd.ExecuteNonQuery()


        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf

        End Try

        Return sResulta

    End Function

    Public Shared Function Mkt_fact3100(ByRef pPgbGlobal As ProgressBar, _
                                             pPgbParcial As ProgressBar, _
                                             plblCurrentOp As Label, _
                                             plblTable As Label,
                                             Optional ByRef pexError As Exception = Nothing) As String

        Dim oCmd As SqlClient.SqlCommand
        Dim sResulta As String = ""

        Try


            '**********************************************************************************************
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Importing 3100 CONTENIDO Free or contenttoolpayment" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "insert into Tsd_transactions (" & _
                                    "sd_source," & _
                                    "processRef, " & _
                                    "TRC_Code," & _
                                    "userId," & _
                                    "created," & _
                                    "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "units, " & _
                                    "source," & _
                                    "sourceId," & _
                                    "productId," & _
                                    "presentationId, " & _
                                    "regularity," & _
                                    "unity, " & _
                                    "observ " & _
                                " ) " & _
                                "select " & _
                                    "sd_source, " & _
                                    "'3100 free contenido', " & _
                                    "'3100' as TRC_Code, " & _
                                    "content_tool_customer_id, " & _
                                    "created_at, " & _
                                    "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "1 as units, " & _
                                    "'FREE' as source, " & _
                                    "increment_id as sourceid, " & _
                                    "0, " & _
                                    "content_tool_product_id," & _
                                    "frequency, " & _
                                    "period,  " & _
                                    "name as observ " & _
                                    "from T_Marketplace " & _
                                    "where  sd_source = 0  " & _
                                    "       and ( " & FREE_CONDITION & " )  " & _
                                    "       and (product_type = 'Presentation') "
            oCmd.ExecuteNonQuery()


        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf

        End Try

        Return sResulta

    End Function


    Public Shared Function Mkt_5000(ByRef pPgbGlobal As ProgressBar, _
                                             pPgbParcial As ProgressBar, _
                                             plblCurrentOp As Label, _
                                             plblTable As Label,
                                             Optional ByRef pexError As Exception = Nothing) As String

        Dim oCmd As SqlClient.SqlCommand
        Dim sResulta As String = ""

        Try


            '**********************************************************************************************
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Importing 5000 Varios" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "insert into Tsd_transactions (" & _
                                    "sd_source," & _
                                    "processRef, " & _
                                    "TRC_Code," & _
                                    "userId," & _
                                    "created," & _
                                     "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "units, " & _
                                    "source," & _
                                    "sourceId," & _
                                    "productId," & _
                                    "presentationId, " & _
                                    "regularity," & _
                                    "unity, " & _
                                    "observ " & _
                                " ) " & _
                                "select " & _
                                    "sd_source, " & _
                                    "'5000 Varios', " & _
                                    "'5000' as TRC_Code, " & _
                                    "content_tool_customer_id, " & _
                                    "created_at, " & _
                                    "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "1 as units, " & _
                                    "'FREE' as source, " & _
                                    "increment_id as sourceid, " & _
                                    "content_tool_product_id, " & _
                                    "0," & _
                                    "frequency, " & _
                                    "period,  " & _
                                    "name as observ " & _
                                    "from T_Marketplace " & _
                                    "where  sd_source = 0  " & _
                                    "       and ( " & FREE_CONDITION & "  ) or (method like '%paypalexpresscheckoutprofilecreated%') " & _
                                    "       and (product_type = 'License') "
            oCmd.ExecuteNonQuery()


        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf

        End Try

        Return sResulta

    End Function

    Public Shared Function Mkt_fact3000(ByRef pPgbGlobal As ProgressBar, _
                                             pPgbParcial As ProgressBar, _
                                             plblCurrentOp As Label, _
                                             plblTable As Label,
                                             Optional ByRef pexError As Exception = Nothing) As String

        Dim oCmd As SqlClient.SqlCommand
        Dim sResulta As String = ""

        Try


            '**********************************************************************************************
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Importing Fact 3000 PAYPAL Contenido" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "insert into Tsd_transactions (" & _
                                    "sd_source," & _
                                    "processRef, " & _
                                    "TRC_Code," & _
                                    "userId," & _
                                    "created," & _
                                    "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "units, " & _
                                    "source," & _
                                    "sourceId," & _
                                    "productId," & _
                                    "presentationId, " & _
                                    "regularity," & _
                                    "unity, " & _
                                    "observ " & _
                                " ) " & _
                                "select " & _
                                    "sd_source, " & _
                                    "'Fact 3000', " & _
                                    "'3000' as TRC_Code, " & _
                                    "content_tool_customer_id, " & _
                                    "created_at, " & _
                                     "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "1 as units, " & _
                                    "'PAYPAL' as source, " & _
                                    "increment_id as sourceid, " & _
                                    "0, " & _
                                    "content_tool_product_id," & _
                                    "frequency, " & _
                                    "period,  " & _
                                    "name as observ " & _
                                    "from T_Marketplace " & _
                                    "where  sd_source = 0  " & _
                                    "       and ( " & PAYPAL_CONDITION & " )  " & _
                                    "       and (product_type = 'Presentation') "
            oCmd.ExecuteNonQuery()


        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf

        End Try

        Return sResulta

    End Function

    Public Shared Function Mkt_fact3300(ByRef pPgbGlobal As ProgressBar, _
                                             pPgbParcial As ProgressBar, _
                                             plblCurrentOp As Label, _
                                             plblTable As Label,
                                             Optional ByRef pexError As Exception = Nothing) As String

        Dim oCmd As SqlClient.SqlCommand
        Dim sResulta As String = ""

        Try


            '**********************************************************************************************
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Importing Fact 3300 MANUAL Contenido" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "insert into Tsd_transactions (" & _
                                    "sd_source," & _
                                    "processRef, " & _
                                    "TRC_Code," & _
                                    "userId," & _
                                    "created," & _
                                     "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "units, " & _
                                    "source," & _
                                    "sourceId," & _
                                    "productId," & _
                                    "presentationId, " & _
                                    "regularity," & _
                                    "unity, " & _
                                    "observ " & _
                                " ) " & _
                                "select " & _
                                    "sd_source, " & _
                                    "'3300 Manual Contenido', " & _
                                    "'3300' as TRC_Code, " & _
                                    "content_tool_customer_id, " & _
                                    "created_at, " & _
                                     "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "1 as units, " & _
                                    "'MANUAL' as source, " & _
                                    "increment_id as sourceid, " & _
                                    "0, " & _
                                    "content_tool_product_id," & _
                                    "frequency, " & _
                                    "period,  " & _
                                    "name as observ " & _
                                    "from T_Marketplace " & _
                                    "where  sd_source = 0  " & _
                                    "       and ( " & MANUAL_CONDITION & " )  " & _
                                    "       and (product_type = 'Presentation') "
            oCmd.ExecuteNonQuery()


        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf

        End Try

        Return sResulta

    End Function

    Public Shared Function Mkt_fact3450(ByRef pPgbGlobal As ProgressBar, _
                                             pPgbParcial As ProgressBar, _
                                             plblCurrentOp As Label, _
                                             plblTable As Label,
                                             Optional ByRef pexError As Exception = Nothing) As String

        Dim oCmd As SqlClient.SqlCommand
        Dim sResulta As String = ""

        Try


            '**********************************************************************************************
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Importing Fact 3450 AYCE Contenido" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "insert into Tsd_transactions (" & _
                                    "sd_source," & _
                                    "processRef, " & _
                                    "TRC_Code," & _
                                    "userId," & _
                                    "created," & _
                                     "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "units, " & _
                                    "source," & _
                                    "sourceId," & _
                                    "productId," & _
                                    "presentationId, " & _
                                    "regularity," & _
                                    "unity, " & _
                                    "observ " & _
                                " ) " & _
                                "select " & _
                                    "sd_source, " & _
                                    "'3450 AYCE Contenido', " & _
                                    "'3450' as TRC_Code, " & _
                                    "content_tool_customer_id, " & _
                                    "created_at, " & _
                                    "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "1 as units, " & _
                                    "'MANUAL' as source, " & _
                                    "increment_id as sourceid, " & _
                                    "0, " & _
                                    "content_tool_product_id," & _
                                    "frequency, " & _
                                    "period,  " & _
                                    "name as observ " & _
                                    "from T_Marketplace " & _
                                    "where  sd_source = 0  " & _
                                    "       and ( " & AYCE_CONDITION & " )  " & _
                                    "       and (product_type = 'Presentation') "
            oCmd.ExecuteNonQuery()


        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf

        End Try

        Return sResulta

    End Function

    Public Shared Function Mkt_fact3050(ByRef pPgbGlobal As ProgressBar, _
                                            pPgbParcial As ProgressBar, _
                                            plblCurrentOp As Label, _
                                            plblTable As Label,
                                            Optional ByRef pexError As Exception = Nothing) As String

        Dim oCmd As SqlClient.SqlCommand
        Dim sResulta As String = ""

        Try


            '**********************************************************************************************
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Importing Fact 3050 IOS Contenido" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "insert into Tsd_transactions (" & _
                                    "sd_source," & _
                                    "processRef, " & _
                                    "TRC_Code," & _
                                    "userId," & _
                                    "created," & _
                                     "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "units, " & _
                                    "source," & _
                                    "sourceId," & _
                                    "productId," & _
                                    "presentationId, " & _
                                    "regularity," & _
                                    "unity, " & _
                                    "observ " & _
                                " ) " & _
                                "select " & _
                                    "sd_source, " & _
                                    "'Fact 3050 IOS Contenido', " & _
                                    "'3050' as TRC_Code, " & _
                                    "content_tool_customer_id, " & _
                                    "created_at, " & _
                                    "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "1 as units, " & _
                                    "'IOS' as source, " & _
                                    "increment_id as sourceid, " & _
                                    "0, " & _
                                    "content_tool_product_id," & _
                                    "frequency, " & _
                                    "period,  " & _
                                    "name as observ " & _
                                    "from T_Marketplace " & _
                                    "where  sd_source = 0  " & _
                                    "       and ( " & IOS_CONDITION & " )  " & _
                                    "       and (product_type = 'Presentation') "
            oCmd.ExecuteNonQuery()


        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf

        End Try

        Return sResulta

    End Function

    Public Shared Function Mkt_fact2050(ByRef pPgbGlobal As ProgressBar, _
                                             pPgbParcial As ProgressBar, _
                                             plblCurrentOp As Label, _
                                             plblTable As Label,
                                             Optional ByRef pexError As Exception = Nothing) As String

        Dim oCmd As SqlClient.SqlCommand
        Dim sResulta As String = ""

        Try


            '**********************************************************************************************
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Importing Fact 2050 IOS" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "insert into Tsd_transactions (" & _
                                    "sd_source," & _
                                    "processRef, " & _
                                    "TRC_Code," & _
                                    "userId," & _
                                    "created," & _
                                     "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "units, " & _
                                    "source," & _
                                    "sourceId," & _
                                    "productId," & _
                                    "presentationId, " & _
                                    "regularity," & _
                                    "unity, " & _
                                    "observ " & _
                                " ) " & _
                                "select " & _
                                    "sd_source, " & _
                                    "'Fact 2050 IOS LICENSE', " & _
                                    "'2050' as TRC_Code, " & _
                                    "content_tool_customer_id, " & _
                                    "created_at, " & _
                                     "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "1 as units, " & _
                                    "'IOS' as source, " & _
                                    "increment_id as sourceid, " & _
                                    "content_tool_product_id, " & _
                                    "0," & _
                                    "frequency, " & _
                                    "period,  " & _
                                    "name as observ " & _
                                    "from T_Marketplace " & _
                                    "where  sd_source = 0  " & _
                                    "       and ( " & IOS_CONDITION & " )  " & _
                                    "       and (product_type = 'License') "
            oCmd.ExecuteNonQuery()


        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf

        End Try

        Return sResulta

    End Function

    Public Shared Function Mkt_fact2400(ByRef pPgbGlobal As ProgressBar, _
                                            pPgbParcial As ProgressBar, _
                                            plblCurrentOp As Label, _
                                            plblTable As Label,
                                            Optional ByRef pexError As Exception = Nothing) As String

        Dim oCmd As SqlClient.SqlCommand
        Dim sResulta As String = ""

        Try


            '**********************************************************************************************
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Importing Fact 2400 Private Library" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "insert into Tsd_transactions (" & _
                                    "sd_source," & _
                                    "processRef, " & _
                                    "TRC_Code," & _
                                    "userId," & _
                                    "created," & _
                                    "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "units, " & _
                                    "source," & _
                                    "sourceId," & _
                                    "productId," & _
                                    "presentationId, " & _
                                    "regularity," & _
                                    "unity, " & _
                                    "observ " & _
                                " ) " & _
                                "select " & _
                                    "sd_source, " & _
                                    "'Fact 2400 Private Library ', " & _
                                    "'2400' as TRC_Code, " & _
                                    "content_tool_customer_id, " & _
                                    "created_at, " & _
                                    "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "1 as units, " & _
                                    "'MANUAL' as source, " & _
                                    "increment_id as sourceid, " & _
                                    "content_tool_product_id, " & _
                                    "0," & _
                                    "frequency, " & _
                                    "period,  " & _
                                    "name as observ " & _
                                    "from T_Marketplace " & _
                                    "where  sd_source = 0  " & _
                                    "       and (product_type = 'Private Library') "
            oCmd.ExecuteNonQuery()


        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf

        End Try

        Return sResulta

    End Function

    Public Shared Function Mkt_fact2350(ByRef pPgbGlobal As ProgressBar, _
                                            pPgbParcial As ProgressBar, _
                                            plblCurrentOp As Label, _
                                            plblTable As Label,
                                            Optional ByRef pexError As Exception = Nothing) As String

        Dim oCmd As SqlClient.SqlCommand
        Dim sResulta As String = ""

        Try


            '**********************************************************************************************
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Importing Fact 2350 AYCE" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "insert into Tsd_transactions (" & _
                                    "sd_source," & _
                                    "processRef, " & _
                                    "TRC_Code," & _
                                    "userId," & _
                                    "created," & _
                                    "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "units, " & _
                                    "source," & _
                                    "sourceId," & _
                                    "productId," & _
                                    "presentationId, " & _
                                    "regularity," & _
                                    "unity, " & _
                                    "observ " & _
                                " ) " & _
                                "select " & _
                                    "sd_source, " & _
                                    "'Fact 2350 AYCE ', " & _
                                    "'2350' as TRC_Code, " & _
                                    "content_tool_customer_id, " & _
                                    "created_at, " & _
                                    "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "1 as units, " & _
                                    "'MANUAL' as source, " & _
                                    "increment_id as sourceid, " & _
                                    "content_tool_product_id, " & _
                                    "0," & _
                                    "frequency, " & _
                                    "period,  " & _
                                    "name as observ " & _
                                    "from T_Marketplace " & _
                                    "where  sd_source = 0  " & _
                                    "       and (product_type = 'AYCE') "
            oCmd.ExecuteNonQuery()


        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf

        End Try

        Return sResulta

    End Function


    Public Shared Function Mkt_fact2000(ByRef pPgbGlobal As ProgressBar, _
                                             pPgbParcial As ProgressBar, _
                                             plblCurrentOp As Label, _
                                             plblTable As Label,
                                             Optional ByRef pexError As Exception = Nothing) As String

        Dim oCmd As SqlClient.SqlCommand
        Dim sResulta As String = ""

        Try


            '**********************************************************************************************
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Importing Fact 2000 PAYPAL" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "insert into Tsd_transactions (" & _
                                    "sd_source," & _
                                    "processRef, " & _
                                    "TRC_Code," & _
                                    "userId," & _
                                    "created," & _
                                     "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "units, " & _
                                    "source," & _
                                    "sourceId," & _
                                    "productId," & _
                                    "presentationId, " & _
                                    "regularity," & _
                                    "unity, " & _
                                    "observ " & _
                                " ) " & _
                                "select " & _
                                    "sd_source, " & _
                                    "'Fact 2000 PAYPAL license', " & _
                                    "'2000' as TRC_Code, " & _
                                    "content_tool_customer_id, " & _
                                    "created_at, " & _
                                     "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "1 as units, " & _
                                    "'PAYPAL' as source, " & _
                                    "increment_id as sourceid, " & _
                                    "content_tool_product_id, " & _
                                    "0," & _
                                    "frequency, " & _
                                    "period,  " & _
                                    "name as observ " & _
                                    "from T_Marketplace " & _
                                    "where  sd_source = 0  " & _
                                    "       and ( " & PAYPAL_CONDITION & " )  " & _
                                    "       and (product_type = 'License') "
            oCmd.ExecuteNonQuery()


        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf

        End Try

        Return sResulta

    End Function

    Public Shared Function Mkt_fact2100(ByRef pPgbGlobal As ProgressBar, _
                                             pPgbParcial As ProgressBar, _
                                             plblCurrentOp As Label, _
                                             plblTable As Label,
                                             Optional ByRef pexError As Exception = Nothing) As String

        Dim oCmd As SqlClient.SqlCommand
        Dim sResulta As String = ""

        Try


            '**********************************************************************************************
            ProgressBarAdd(pPgbGlobal)
            oCmd = goConNear.CreateCommand
            oCmd.CommandTimeout = 999999
            plblCurrentOp.Text = "Importing Fact 2100 MANUAL" : Application.DoEvents()
            sResulta += "Processing: " & plblCurrentOp.Text & " at " & Now.ToString & vbCrLf
            oCmd.CommandText = "insert into Tsd_transactions (" & _
                                    "sd_source," & _
                                    "processRef, " & _
                                    "TRC_Code," & _
                                    "userId," & _
                                    "created," & _
                                     "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "units, " & _
                                    "source," & _
                                    "sourceId," & _
                                    "productId," & _
                                    "presentationId, " & _
                                    "regularity," & _
                                    "unity, " & _
                                    "observ " & _
                                " ) " & _
                                "select " & _
                                    "sd_source, " & _
                                    "'Fact 2100 PAYPAL license', " & _
                                    "'2100' as TRC_Code, " & _
                                    "content_tool_customer_id, " & _
                                    "created_at, " & _
                                     "price," & _
                                    "grand_total," & _
                                    "subtotal," & _
                                    "discount_amount," & _
                                    "coupon_rule_id," & _
                                    "coupon_rule_name," & _
                                    "coupon_code," & _
                                    "1 as units, " & _
                                    "'MANUAL' as source, " & _
                                    "increment_id as sourceid, " & _
                                    "content_tool_product_id, " & _
                                    "0," & _
                                    "frequency, " & _
                                    "period,  " & _
                                    "name as observ " & _
                                    "from T_Marketplace " & _
                                    "where  sd_source = 0  " & _
                                    "       and ( " & MANUAL_CONDITION & " )  " & _
                                    "       and (product_type = 'License') "
            oCmd.ExecuteNonQuery()


        Catch ex As Exception
            pexError = ex
            sResulta += ex.ToString & vbCrLf

        End Try

        Return sResulta

    End Function

    Public Shared Function ArmoCadena(ByRef a() As String)
        Dim cadena As String = ""
        For i = 0 To UBound(a)
            If (i > 0) Then
                cadena = cadena & ","
            End If
            cadena = cadena + a(i)
        Next

        Return "(" & cadena & ")"

    End Function




End Class



