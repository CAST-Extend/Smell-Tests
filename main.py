#import cast_upgrade_1_5_21  #@UnusedImport
import os
import time
import xlsxwriter
from cast.application import ApplicationLevelExtension
from tkinter.font import BOLD


class Report(ApplicationLevelExtension):
    
    def start_application(self, application):
        """
        Called before analysis.
        
        .. versionadded:: CAIP 8.3
        
        :type application: :class:`cast.application.Application`
        @type application: cast.application.Application
        """
        pass
    
    
    def end_application(self, application):
        """
        Called at the end of application's analysis.

        :type application: :class:`cast.application.Application`
        @type application: cast.application.Application
        """
        pass

    def after_module(self, application):
        """
        Called after module content creation.
        
        .. versionadded:: CAIP 8.3
        
        :type application: :class:`cast.application.Application`
        @type application: cast.application.Application
        """
        
        pass
        

    def after_snapshot(self, application):
        """
        Called after module content creation.
        Gives you the central's application.
        
        .. versionadded:: CAIP 8.3
        
        :type application: :class:`cast.application.central.Application`        
        @type application: cast.application.central.Application
        """
        
        # do your things afte rmodule...
        # this import may fail in versions < 8.3
        from cast.application import publish_report # @UnresolvedImport
        
        # generate a path in LISA my_report<timestamp>.xlxs 
        report_path = os.path.join(self.get_plugin().intermediate, time.strftime("my_report%Y%m%d_%H%M%S.xlsx"))
        
        # calculate and fill file content
        
        # we use XlsxWriter to write excel 
        # @see https://xlsxwriter.readthedocs.io/
        workbook = xlsxwriter.Workbook(report_path)
        # @type workbook: xlsxwriter.Workbook
        
        worksheet = workbook.add_worksheet('Smell Tests')
        # @type worksheet: xlsxwriter.Worksheet
        
        
        # kb represent a local
        #kb = application.get_knowledge_base()
        kb = application.get_application_configuration().get_analysis_service()
        central = application.get_central() 
        # @type kb : cast.application.KnowledgeBase
        
        # we count elements in table keys
        total_count = 0
        covered_count = 0
        percent_coverage = ''
        RAGCoverage = ''
        total_txns = ''
        empty_txns = ''
        percent_empty = ''
        RAGEmptyTxns = ''
        RAGCalibration = ''
        newDLM = ''
        TotalDLM = ''
        InvalidReviwed = ''
        ValidReviewed = ''
        percentageNew = ''
        RAGDLM = ''
        TFP = ''
        DFP = ''
        percentageTFPDFP = ''
        RAGTFPDFP = ''
        Centralcompute_vaue1 = ''
        Centralcompute_vaue2 = ''  
        Centralpercent_coverage = ''
        CentralRAGCoverage = ''
        
        for lineCentral in central.execute_query("""select cast(Programs_Class as text) as compute_value1, cast(FP as text) as compute_value2, cast(fp_to_pgmclass_ratio as text) as percentage_ratio,
case when fp_to_pgmclass_ratio >=0 and fp_to_pgmclass_ratio < 1 then 'RED'
when fp_to_pgmclass_ratio >=1 and fp_to_pgmclass_ratio < 2 then 'AMBER'
when fp_to_pgmclass_ratio >=2 and fp_to_pgmclass_ratio < 4 then 'GREEN'
when fp_to_pgmclass_ratio >=4 and fp_to_pgmclass_ratio < 5 then 'AMBER'
when fp_to_pgmclass_ratio >=5 then 'RED'
end as rag
from(
select pgm_classes.pgm_count Programs_Class, fp_count.fp FP, 
round(
                CASE
                    WHEN pgm_classes.pgm_count = 0::numeric THEN 0::numeric
                    ELSE fp_count.fp / pgm_classes.pgm_count
                END, 2) AS fp_to_pgmclass_ratio  
 from 
(select sum(metric_num_value) as pgm_count
from 
dss_metric_results t1,
dss_metric_types t2, 
dss_objects t3
where
t1.metric_id = t2.metric_id
and t1.object_id = t3.object_id
and t3.object_type_id = -102 
and t1.snapshot_id = (select max(snapshot_id) from dss_snapshots)
and t1.metric_id in ( 10155, 10156)) as pgm_classes,
(select sum(metric_num_value) as fp from 
dss_metric_results t1,
dss_metric_types t2, 
dss_objects t3
where
t1.metric_id = t2.metric_id
and t1.object_id = t3.object_id
and t3.object_type_id = -102 
and t1.snapshot_id = (select max(snapshot_id) from dss_snapshots)
and t1.metric_id in ( 10203, 10204)) as fp_count 
) as dataTable"""):
            Centralcompute_vaue1 = lineCentral[0]
            Centralcompute_vaue2 = lineCentral[1]  
            Centralpercent_coverage = lineCentral[2] 
            CentralRAGCoverage = lineCentral[3]
        
        for line in kb.execute_query("""select total_count,covered_count,cast((( case when covered_count = 0 then 1 else covered_count end)*100/(case when total_count = 0 then 1 else total_count end))as numeric)||'%' as percent_coverage,
CASE when cast(((case when covered_count = 0 then 1 else covered_count end)*100/(case when total_count = 0 then 1 else total_count end))as integer) <= 30 THEN 'RED'
when (cast(((case when covered_count = 0 then 1 else covered_count end)*100/(case when total_count = 0 then 1 else total_count end))as integer) <= 50
and cast(((case when covered_count = 0 then 1 else covered_count end)*100/(case when total_count = 0 then 1 else total_count end))as integer) > 30) THEN 'AMBER'
when cast(((case when covered_count = 0 then 1 else covered_count end)*100/(case when total_count = 0 then 1 else total_count end))as integer) > 50 THEN 'GREEN' end as case
from
(
select count(distinct cdt.object_id) as total_count from cdt_objects cdt, ctt_object_applications ctt where cdt.object_id=ctt.object_id and ctt.properties<>1 and
cdt.object_type_str in  ('Progress Program','Cobol Program','C++ Class','VB.NET Class','ColdFusion Fuse Action','VB MDI Form','SHELL Program','ColdFusion Template','C/C++ File','C# Class','Java Class','VB Module','Session')) total,
(
select count(distinct cdt.object_id) as covered_count  from cdt_objects cdt, ctt_object_applications ctt where cdt.object_id=ctt.object_id and ctt.properties<>1 and ctt.object_id in
 
(select object_id from cdt_objects where object_id in (
select child_id from dss_transactiondetails)
 
union
select parent_id from ctt_object_parents where object_id in (
select child_id from dss_transactiondetails)
 
union
 
select called_id from ctv_links where caller_id in
 (select child_id from dss_transactiondetails)
 union
 select called_id from ctv_links where caller_id in(
select called_id from ctv_links where caller_id in
 (select child_id from dss_transactiondetails))
 
union
select parent_id from ctt_object_parents where object_id in (
select called_id from ctv_links where caller_id in (select child_id from dss_transactiondetails))

union
select parent_id from ctt_object_parents where object_id in (
select called_id from ctv_links where caller_id in(
select called_id from ctv_links where caller_id in
 (select child_id from dss_transactiondetails)))

union
select object_id from ctt_object_parents where parent_id in (
select parent_id from ctt_object_parents where object_id in (
select child_id from dss_transactiondetails))

union
select object_id from ctt_object_parents where parent_id in (
select parent_id from ctt_object_parents where object_id in (
select called_id from ctv_links where caller_id in (select child_id from dss_transactiondetails)))

union
select object_id from ctt_object_parents where parent_id in (
select parent_id from ctt_object_parents where object_id in (
select called_id from ctv_links where caller_id in(
select called_id from ctv_links where caller_id in
 (select child_id from dss_transactiondetails)))))
and cdt.object_type_str in  ('Progress Program','Cobol Program','C++ Class','VB.NET Class','ColdFusion Fuse Action','VB MDI Form','SHELL Program','ColdFusion Template','C/C++ File','C# Class','Java Class','VB Module','Session'))covered"""):
            total_count = line[0]
            covered_count = line[1]  
            percent_coverage = line[2] 
            RAGCoverage = line[3]   
            
        for line2 in kb.execute_query("""select cast((case when tb_all_tr.tr_count = 0 then 1 else tb_all_tr.tr_count end) as text) as compute_value2, cast((case when tb_empty_tr.tr_count = 0 then 1 else tb_empty_tr.tr_count end) as text) as compute_value1 , 
cast( (((case when tb_empty_tr.tr_count = 0 then 1 else tb_empty_tr.tr_count end)  * 100)/ (case when tb_all_tr.tr_count = 0 then 1 else tb_all_tr.tr_count end)) as text) as percentage_ratio,
case when ((((case when tb_empty_tr.tr_count = 0 then 1 else tb_empty_tr.tr_count end)  * 100)/ (case when tb_all_tr.tr_count = 0 then 1 else tb_all_tr.tr_count end)) >= 0 and (((case when tb_empty_tr.tr_count = 0 then 1 else tb_empty_tr.tr_count end)  * 100)/ (case when tb_all_tr.tr_count = 0 then 1 else tb_all_tr.tr_count end)) < 10) then 'GREEN'
when ((((case when tb_empty_tr.tr_count = 0 then 1 else tb_empty_tr.tr_count end)  * 100)/ (case when tb_all_tr.tr_count = 0 then 1 else tb_all_tr.tr_count end)) >=10 and (((case when tb_empty_tr.tr_count = 0 then 1 else tb_empty_tr.tr_count end)  * 100)/ (case when tb_all_tr.tr_count = 0 then 1 else tb_all_tr.tr_count end)) < 30) then 'AMBER'
else 'RED'
end as rag
from
(select count(cob.object_name) as tr_count
from dss_transaction dtr, cdt_objects cob
where dtr.form_id = cob.object_id
and dtr.cal_mergeroot_id = 0  
and dtr.cal_flags not in (  8, 10, 126, 128,136, 138, 256, 258 ) 
) tb_all_tr,
(select count(cob.object_name) as tr_count from dss_transaction dtr, cdt_objects cob 
where dtr.form_id = cob.object_id
and dtr.cal_mergeroot_id = 0  
and dtr.cal_flags not in (  8, 10, 126, 128,136, 138, 256, 258 )  
and DTR.tf_ex=0  ) tb_empty_tr;"""):
            total_txns = line2[0]
            empty_txns = line2[1]  
            percent_empty = line2[2] 
            RAGEmptyTxns = line2[3]   
        
        for line3 in kb.execute_query("""Select cast(count(*)as text) as compute_value1, case when (select count(*) from sys_package_version where package_name like '%com.castsoftware.uc.transactioncalibrationkit%') > 0 then 'GREEN'  else 'RED' end as rag
from sys_package_version where package_name like '%com.castsoftware.uc.transactioncalibrationkit%'"""): 
            RAGCalibration = line3[1]  
            
        for line4 in kb.execute_query("""select cast((select count(*) from acc where prop = 1) as text)  as NewDLM, cast((select count(*) from acc where prop in (1,65537)) as text) as total, cast((select count(*) from acc where prop = 65537) as text) as InvalidReviewed, cast((select count(*) from acc where prop = 32769) as text) as ValidReviewed, 
 cast( (((select count(*) from acc where prop = 1)   * 100 + .000000000001)/ ((select count(*) from acc where prop in (1,65537)) + .000000000001)) as text) as percentage_ratio, 
case when (select count(*) from acc where prop = 1) > 0 then 'RED'  else 'GREEN' end as rag """): 
            newDLM = line4[0]
            TotalDLM = line4[1]
            InvalidReviwed = line4[2]
            ValidReviewed = line4[3]
            percentageNew = line4[4]
            RAGDLM = line4[5]
         
        for line5 in kb.execute_query("""select cast(TFP.TFP as text) as TFP, cast(DFP.DFP as text) as DFP, cast((TFP + .0001)/(DFP +.0001) as text) as percentage_ratio, 
case when (TFP+ .0001)/(DFP+ .0001) >=0 and (TFP+ .0001)/(DFP+ .0001) < 1 then 'RED'
when (TFP+ .0001)/(DFP+ .0001) >=1 and (TFP+ .0001)/(DFP+ .0001) < 2 then 'AMBER'
when (TFP+ .0001)/(DFP+ .0001) >=2 and (TFP+ .0001)/(DFP+ .0001) < 4 then 'GREEN'
when (TFP+ .0001)/(DFP+ .0001) >=4 and (TFP+ .0001)/(DFP+ .0001) < 5 then 'AMBER'
when (TFP+ .0001)/(DFP+ .0001) >=5 then 'RED'
end as rag
from (
select sum(DTR.tf_ex) as TFP
from dss_transaction dtr, cdt_objects cob
where dtr.form_id = cob.object_id
and dtr.cal_mergeroot_id = 0 
and dtr.cal_flags not in ( 8, 10, 126, 128,136, 138, 256, 258 ) 
) as TFP,
( 
select sum(dtf.ilf_ex) as DFP
from dss_datafunction dtf, cdt_objects cob
where dtf.maintable_id = cob.object_id 
and dtf.cal_flags in (0,2)
) as DFP; """): 
            TFP = line5[0]
            DFP = line5[1]
            percentageTFPDFP = line5[2]
            RAGTFPDFP = line5[3] 
            
        # write something at line 0 column 0
        
        
        
        worksheet.write(0, 0, 'Smell Test Checks')
        worksheet.write(0, 1, 'Compute Value 1')
        worksheet.write(0, 2, 'Compute Value 2')
        worksheet.write(0, 3, 'Percent (Ratio)')
        worksheet.write(0, 4, 'Current RAG')
        
        
        worksheet.write(1, 0, 'Artifact Coverage')
        worksheet.write(1, 1, total_count)
        worksheet.write(1, 2, covered_count)
        worksheet.write(1, 3, percent_coverage)
        worksheet.write(1, 4, RAGCoverage)
        
        worksheet.write(2, 0, 'Empty Transactions')
        worksheet.write(2, 1, total_txns)
        worksheet.write(2, 2, empty_txns)
        worksheet.write(2, 3, percent_empty)
        worksheet.write(2, 4, RAGEmptyTxns)
        
        worksheet.write(3, 0, 'Calibration Kit Applied')
        worksheet.write(3, 1, '')
        worksheet.write(3, 2, '')
        worksheet.write(3, 3, '')
        worksheet.write(3, 4, RAGCalibration)
        
        worksheet.write(4, 0, 'DLMs reviewed')
        worksheet.write(4, 1, newDLM)
        worksheet.write(4, 2, TotalDLM)
        worksheet.write(4, 3, InvalidReviwed)
        worksheet.write(4, 4, RAGDLM)
        
        worksheet.write(5, 0, 'TFP DFP Ratio')
        worksheet.write(5, 1, TFP)
        worksheet.write(5, 2, DFP)
        worksheet.write(5, 3, percentageTFPDFP)
        worksheet.write(5, 4, RAGTFPDFP)
                
        worksheet.write(6, 0, 'FP Class Ratio')
        worksheet.write(6, 1, Centralcompute_vaue1)
        worksheet.write(6, 2, Centralcompute_vaue2)
        worksheet.write(6, 3, Centralpercent_coverage)
        worksheet.write(6, 4, CentralRAGCoverage)
        
        
        worksheet1 = workbook.add_worksheet('Extensions')
        worksheet1.write(0, 0, 'EXTENSIONS')
        indexx = 1
        for line6 in kb.execute_query("""select * from sys_package_version where package_name like '\/%'"""): 
            worksheet1.write(indexx, 0, line6[0])
            indexx = indexx +1
        # ...
        
        # we have finished
        workbook.close()
        
        # calculate a status
        # this is a sample : choose one
        
        #status = "Warning"
        #status = "KO"
        status = "Warning"
        
        if RAGCoverage == "GREEN" and RAGEmptyTxns == "GREEN" and RAGCalibration == "GREEN" and RAGDLM =="GREEN" and RAGTFPDFP == "GREEN":
            status = "OK"
        
        # publish report :  
        publish_report('Smell Tests Results', 
                       status, "Onboarding completeness", '', detail_report_path=report_path)
