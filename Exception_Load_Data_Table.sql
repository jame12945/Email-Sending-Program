create or replace PACKAGE BODY CRM_RDT_EXCEPTION_PKG AS

PROCEDURE get_t_rdt_exc_report_022  AS
    rdt_exists NUMBER;
BEGIN
   --SELECT COUNT(*)
   --INTO rdt_exists
   --FROM T_RDT_EXC_REPORT_022
   --WHERE DATA_DATE = i_char_as_of_date;

   --IF (rdt_exists > 0 ) and i_char_as_of_date is not null THEN

        --DELETE  FROM T_RDT_EXC_REPORT_022 WHERE DATA_DATE = i_char_as_of_date;
        execute immediate ' truncate table T_RDT_EXC_REPORT_022 ';
        
        insert into T_RDT_EXC_REPORT_022
        select  DISTINCT t.DATA_DATE,
                t.ACCT_NR,
                t.B_DATA_ORIGIN as SYSTEM_NAME,
                t.CUST_NO as CUST_NR,
                t.ACCT_TYPE,
                t.ACCT_PROD,
                t.ACCOUNT_PURPOSE as BOT_PURPOSE_CODE,
                t.ISIC_CODE,
                t1.RM_CUST_NR,
                t1.bos_cust_id as BOS_CUST_ID,
                trim(CONCAT(CONCAT(COALESCE(CONCAT(
                        CASE 
                            WHEN t1.ASSOCIATION_TYPE = 'J' THEN 'บัญชีร่วม'
                            WHEN t1.ASSOCIATION_TYPE = 'U' THEN 'กิจการร่วมค้าแบบไม่จดทะเบียน'
                            WHEN t1.ASSOCIATION_TYPE = 'C' THEN 'กิจการค้าร่วม'
                            WHEN t1.title_thai = 'อื่นๆ' THEN ''
                            else t1.title_thai
                        END       
                        ,' '),null),
                        COALESCE(CONCAT(t1.cust_name_thai,' '),null)),
                        COALESCE(t1.cust_last_name_thai,null)))  AS customer_name,                    
                t3.bu_code as bu_code,
                DECODE(trim(t4.BU_CODE),'NWM',t4.div_code,'SW',t4.div_code,'SCM',t4.div_code,'SBM',t4.div_code,t4.bc_code) AS  div_bc_code,
                DECODE(trim(t4.BU_CODE),'NWM',t5.div_thai_name,'SW',t5.div_thai_name,'SCM',t5.div_thai_name,'SBM',t5.div_thai_name,t6.bc_thai_name) AS div_bc_name,
                t3.ro_arm_ao_alloc,
                t3.ro_arm_ao_name,
                t3.RM_AM_ALLOC,
                t3.RM_AM_NAME ,
                t3.m_au_tl_name
        FROM  temp_rdt_exc_report_003_022 t,
              SND_EM_CUST_DETAIL t1,
              snd_em_cust_detail_additional t2,
              SN_DM_PORTFOLIO_HIERARCHY t3,
              t_cis_authorized_user t4,
              t_bbl_division_new t5,
              t_bbl_business_center t6 
        where trim(t.rm_cust_nr) = t1.rm_cust_nr
        and t1.bos_cust_id = t2.bos_cust_id
        and t1.bos_cust_id = t3.cust_id
        AND t3.ro_arm_ao_alloc = t4.allocation_code and t4.end_date is null
        and t4.div_code = t5.div_code(+)
        and t4.bc_code = t6.bc_code(+)
        and t.EXCEPTION_CODE = '022';
   --END IF;
      COMMIT;  
END get_t_rdt_exc_report_022;

PROCEDURE get_t_rdt_exc_report_003  AS
    rdt_exists NUMBER;
BEGIN
   --SELECT COUNT(*)
   --INTO rdt_exists
   --FROM T_RDT_EXC_REPORT_003
   --WHERE DATA_DATE = i_char_as_of_date;

   --IF (rdt_exists > 0 ) and i_char_as_of_date is not null THEN

        --DELETE  FROM T_RDT_EXC_REPORT_003 WHERE DATA_DATE = i_char_as_of_date;
        execute immediate ' truncate table T_RDT_EXC_REPORT_003 ';
        
        insert into T_RDT_EXC_REPORT_003
        select  DISTINCT t.DATA_DATE,
                t.ACCT_NR,
                t.B_DATA_ORIGIN as SYSTEM_NAME,
                t.CUST_NO as CUST_NR,
                t.ACCT_TYPE,
                t.ACCT_PROD,
                t.ACCOUNT_PURPOSE as BOT_PURPOSE_CODE,
                t.ISIC_CODE,
                t1.RM_CUST_NR,
                t1.bos_cust_id as BOS_CUST_ID,
                trim(CONCAT(CONCAT(COALESCE(CONCAT(
                        CASE 
                            WHEN t1.ASSOCIATION_TYPE = 'J' THEN 'บัญชีร่วม'
                            WHEN t1.ASSOCIATION_TYPE = 'U' THEN 'กิจการร่วมค้าแบบไม่จดทะเบียน'
                            WHEN t1.ASSOCIATION_TYPE = 'C' THEN 'กิจการค้าร่วม'
                            WHEN t1.title_thai = 'อื่นๆ' THEN ''
                            else t1.title_thai
                        END       
                        ,' '),null),
                        COALESCE(CONCAT(t1.cust_name_thai,' '),null)),
                        COALESCE(t1.cust_last_name_thai,null)))  AS customer_name,                    
                t3.bu_code as bu_code,
                DECODE(trim(t4.BU_CODE),'NWM',t4.div_code,'SW',t4.div_code,'SCM',t4.div_code,'SBM',t4.div_code,t4.bc_code) AS  div_bc_code,
                DECODE(trim(t4.BU_CODE),'NWM',t5.div_thai_name,'SW',t5.div_thai_name,'SCM',t5.div_thai_name,'SBM',t5.div_thai_name,t6.bc_thai_name) AS div_bc_name,
                t3.ro_arm_ao_alloc,
                t3.ro_arm_ao_name,
                t3.RM_AM_ALLOC,
                t3.RM_AM_NAME ,
                t3.m_au_tl_name
        FROM  temp_rdt_exc_report_003_022 t,
              SND_EM_CUST_DETAIL t1,
              snd_em_cust_detail_additional t2,
              SN_DM_PORTFOLIO_HIERARCHY t3,
              t_cis_authorized_user t4,
              t_bbl_division_new t5,
              t_bbl_business_center t6 
        where trim(t.rm_cust_nr) = t1.rm_cust_nr
        and t1.bos_cust_id = t2.bos_cust_id
        and t1.bos_cust_id = t3.cust_id
        AND t3.ro_arm_ao_alloc = t4.allocation_code and t4.end_date is null
        and t4.div_code = t5.div_code(+)
        and t4.bc_code = t6.bc_code(+)
        and t.EXCEPTION_CODE = '003';
   --END IF;
     COMMIT;  
END get_t_rdt_exc_report_003;

  PROCEDURE get_t_rdt_exc_report_011  AS
      rdt_exists NUMBER;
  BEGIN
     --Step 1: Check t_rdt_exc_report_011 existing data
                    /* SELECT COUNT(*)
                      INTO rdt_exists
                      FROM t_rdt_exc_report_011
                      WHERE DATA_DATE = i_char_as_of_date;

                       IF (rdt_exists > 0 ) and i_char_as_of_date is not null THEN

                       DELETE  FROM t_rdt_exc_report_011 WHERE DATA_DATE = i_char_as_of_date;

                       INSERT INTO t_rdt_exc_report_011 (
                                                           DATA_DATE,
                                                           LINKAGE_NO,
                                                           COLLATERAL_TYPE,
                                                           COLLATERAL_TYPE_DESC,
                                                           COLLATERAL_SUB_TYPE1,
                                                           STOCK_CODE,
                                                           APPRASIAL_VALUE,
                                                           APPRASIAL_VALUE_CURRENCY,
                                                           APPRASIAL_DATE,
                                                           ISSUER_NAME,
                                                           RM_CUST_NR,
                                                           BOS_CUST_ID,
                                                           CUST_NAME,
                                                           BU_CODE,
                                                           DIV_BC_CODE,
                                                           DIV_BC_NAME,
                                                           RO_ARM_AO_ALLOC_CODE,
                                                           RO_ARM_AO_NAME,
                                                           RM_AM_ALLOCA_CODE,
                                                           RM_AM_NAME,
                                                           M_TL_NAME
                                                        )
                    SELECT  DISTINCT
                    t1.data_date,
                    t1.linkage_no as cms_linkage_no,
                    t1.collateral_type,
                    t1.collateral_desc,
                    t1.collateral_sub_type1,
                    t1.stock_code,
                    t1.total_out_value as appraisal_value,
                    t1.currency_app_out as appraisal_value_currency,
                    t1.outsource_date as appraisal_date,
                    t1.issuer_name,
                    t3.rm_cust_nr as RM_Number,
                    t3.cust_id AS bos_cutomer_id,
                     trim(CONCAT(CONCAT(COALESCE(CONCAT(
                    CASE 
                        WHEN t3.ASSOCIATION_TYPE = 'J' THEN 'บัญชีร่วม'
                        WHEN t3.ASSOCIATION_TYPE = 'U' THEN 'กิจการร่วมค้าแบบไม่จดทะเบียน'
                        WHEN t3.ASSOCIATION_TYPE = 'C' THEN 'กิจการค้าร่วม'
                        WHEN t3.cust_title_thai_desc = 'อื่นๆ' THEN ''
                        else t3.cust_title_thai_desc
                    END       
                    ,' '),null),
                    COALESCE(CONCAT(t3.cust_thai_name,' '),null)),
                    COALESCE(t3.cust_thai_last_name,null)))  AS customer_name,  
                    t4.bu_code as bu_code,
                    DECODE(trim(t7.BU_CODE),'NWM',t7.div_code,'SW',t7.div_code,'SCM',t7.div_code,'SBM',t7.div_code,t7.bc_code) AS  div_bc_code,
                    DECODE(trim(t7.BU_CODE),'NWM',t6.div_thai_name,'SW',t6.div_thai_name,'SCM',t6.div_thai_name,'SBM',t6.div_thai_name,t8.bc_thai_name) AS div_bc_name,
                    t4.ro_arm_ao_alloc,
                    t4.ro_arm_ao_name,
                    t4.RM_AM_ALLOC,
                    t4.RM_AM_NAME ,
                    t4.m_au_tl_name
                    FROM (SELECT * FROM colloan.t_cms_acl_stock WHERE data_date = (SELECT MAX(data_date)FROM colloan.t_cms_acl_stock))t1
                    LEFT JOIN t_cms_collateral_relation t2 ON t2.linkage_no = t1.linkage_no and t2.as_of_date = t1.data_date
                    LEFT JOIN V_SND_EM_CUST_DETAIL t3 ON t3.cust_id = t2.cust_id
                    LEFT JOIN SN_DM_PORTFOLIO_HIERARCHY t4 on t4.cust_id = t3.cust_id 
                    LEFT JOIN t_bbl_division_new t6 ON t3.div_code = t6.div_code AND t3.bu_code = t6.bu_code
                    LEFT JOIN t_cis_authorized_user t7 ON t7.allocation_code = t4.ro_arm_ao_alloc and t7.end_date is null
                    LEFT JOIN t_bbl_business_center t8 ON t7.bc_code = t8.bc_code
                    WHERE  t1.COLLATERAL_TYPE = '04' and t1.COLLATERAL_SUB_TYPE1 IN (3,5) and t1.OUTSOURCE_DATE is null 
                    and t1.data_date = (SELECt MAX(t2.data_date) FROM  colloan.t_cms_acl_stock t2)  and t1.data_date = i_char_as_of_date
                    and t3.status is not null and t7.div_code  = t6.div_code
                    order by t1.DATA_DATE desc;

                   END IF;

IF i_char_as_of_date is null THEN  */ 
--DELETE  FROM t_rdt_exc_report_011 where data_date = i_char_as_of_date;
execute immediate ' truncate table T_RDT_EXC_REPORT_011 ';

INSERT INTO t_rdt_exc_report_011 (
    DATA_DATE,
    LINKAGE_NO,
    COLLATERAL_TYPE,
    COLLATERAL_TYPE_DESC,
    COLLATERAL_SUB_TYPE1,
    STOCK_CODE,
    APPRASIAL_VALUE,
    APPRASIAL_VALUE_CURRENCY,
    APPRASIAL_DATE,
    ISSUER_NAME,
    RM_CUST_NR,
    BOS_CUST_ID,
    CUST_NAME,
    BU_CODE,
    DIV_BC_CODE,
    DIV_BC_NAME,
    RO_ARM_AO_ALLOC_CODE,
    RO_ARM_AO_NAME,
    RM_AM_ALLOCA_CODE,
    RM_AM_NAME,
    M_TL_NAME
)
SELECT  DISTINCT
                    t1.data_date,
                    t1.linkage_no as cms_linkage_no,
                    t1.collateral_type,
                    t1.collateral_desc,
                    t1.collateral_sub_type1,
                    t1.stock_code,
                    t1.total_out_value as appraisal_value,
                    t1.currency_app_out as appraisal_value_currency,
                    t1.outsource_date as appraisal_date,
                    t1.issuer_name,
                    t3.rm_cust_nr as RM_Number,
                    t3.cust_id AS bos_cutomer_id,
                    --CONCAT(CONCAT(COALESCE(CONCAT(t3.cust_title_thai_desc,' '),null),COALESCE(CONCAT(t3.cust_thai_name,' '),null)),COALESCE(t3.cust_thai_last_name,null))  AS customer_name,
                     trim(CONCAT(CONCAT(COALESCE(CONCAT(
                    CASE 
                        WHEN t3.ASSOCIATION_TYPE = 'J' THEN 'บัญชีร่วม'
                        WHEN t3.ASSOCIATION_TYPE = 'U' THEN 'กิจการร่วมค้าแบบไม่จดทะเบียน'
                        WHEN t3.ASSOCIATION_TYPE = 'C' THEN 'กิจการค้าร่วม'
                        WHEN t3.cust_title_thai_desc = 'อื่นๆ' THEN ''
                        else t3.cust_title_thai_desc
                    END       
                    ,' '),null),
                    COALESCE(CONCAT(t3.cust_thai_name,' '),null)),
                    COALESCE(t3.cust_thai_last_name,null)))  AS customer_name,  
                    t4.bu_code as bu_code,
                    DECODE(trim(t7.BU_CODE),'NWM',t7.div_code,'SW',t7.div_code,'SCM',t7.div_code,'SBM',t7.div_code,t7.bc_code) AS  div_bc_code,
                    DECODE(trim(t7.BU_CODE),'NWM',t6.div_thai_name,'SW',t6.div_thai_name,'SCM',t6.div_thai_name,'SBM',t6.div_thai_name,t8.bc_thai_name) AS div_bc_name,
                    t4.ro_arm_ao_alloc,
                    t4.ro_arm_ao_name,
                    t4.RM_AM_ALLOC,
                    t4.RM_AM_NAME ,
                    t4.m_au_tl_name
                    FROM (SELECT * FROM colloan.t_cms_acl_stock WHERE data_date = (SELECT MAX(data_date)FROM colloan.t_cms_acl_stock))t1
                    LEFT JOIN t_cms_collateral_relation t2 ON t2.linkage_no = t1.linkage_no and t2.as_of_date = t1.data_date
                    LEFT JOIN V_SND_EM_CUST_DETAIL t3 ON t3.cust_id = t2.cust_id
                    LEFT JOIN SN_DM_PORTFOLIO_HIERARCHY t4 on t4.cust_id = t3.cust_id 
                    LEFT JOIN t_bbl_division_new t6 ON t3.div_code = t6.div_code AND t3.bu_code = t6.bu_code
                    LEFT JOIN t_cis_authorized_user t7 ON t7.allocation_code = t4.ro_arm_ao_alloc and t7.end_date is null
                    LEFT JOIN t_bbl_business_center t8 ON t7.bc_code = t8.bc_code
                    WHERE  t1.COLLATERAL_TYPE = '04' and t1.COLLATERAL_SUB_TYPE1 IN (3,5) and t1.OUTSOURCE_DATE is null 
                    and t1.data_date = (SELECt MAX(t2.data_date) FROM  colloan.t_cms_acl_stock t2)
                    and t3.status is not null and t7.div_code  = t6.div_code
                    order by t1.DATA_DATE desc;
    --END IF;
    COMMIT;
  END get_t_rdt_exc_report_011;


  PROCEDURE get_t_rdt_exc_report_012  AS
                                     rdt_exists NUMBER;
    BEGIN
                    /*
                    --Step 1: Check t_rdt_exc_report_012 existing data
                      SELECT COUNT(*)
                      INTO rdt_exists
                      FROM t_rdt_exc_report_012
                      WHERE DATA_DATE = i_char_as_of_date;
                    --Step 2: If have duplicate data_date delete them else it will insert data 
                    IF (rdt_exists > 0) and i_char_as_of_date is not null THEN
                     INSERT INTO t_rdt_exc_report_012 (
                                                           DATA_DATE,
                                                           LINKAGE_NO,
                                                           COLLATERAL_TYPE,
                                                           COLLATERAL_TYPE_DESC,
                                                           COLLATERAL_SUB_TYPE1,
                                                           STOCK_CODE,
                                                           APPRASIAL_VALUE,
                                                           APPRASIAL_VALUE_CURRENCY,
                                                           APPRASIAL_DATE,
                                                           ISSUER_NAME,
                                                           RM_CUST_NR,
                                                           BOS_CUST_ID,
                                                           CUST_NAME,
                                                           BU_CODE,
                                                           DIV_BC_CODE,
                                                           DIV_BC_NAME,
                                                           RO_ARM_AO_ALLOC_CODE,
                                                           RO_ARM_AO_NAME,
                                                           RM_AM_ALLOCA_CODE,
                                                           RM_AM_NAME,
                                                           M_TL_NAME
                                                        )
                    SELECT  DISTINCT
                    t1.data_date,
                    t1.linkage_no as cms_linkage_no,
                    t1.collateral_type,
                    t1.collateral_desc,
                    t1.collateral_sub_type1,
                    t1.stock_code,
                    t1.total_out_value as appraisal_value,
                    t1.currency_app_out as appraisal_value_currency,
                    t1.outsource_date as appraisal_date,
                    t1.issuer_name,
                    t3.rm_cust_nr as RM_Number,
                    t3.cust_id AS bos_cutomer_id,
                     trim(CONCAT(CONCAT(COALESCE(CONCAT(
                    CASE 
                        WHEN t3.ASSOCIATION_TYPE = 'J' THEN 'บัญชีร่วม'
                        WHEN t3.ASSOCIATION_TYPE = 'U' THEN 'กิจการร่วมค้าแบบไม่จดทะเบียน'
                        WHEN t3.ASSOCIATION_TYPE = 'C' THEN 'กิจการค้าร่วม'
                        WHEN t3.cust_title_thai_desc = 'อื่นๆ' THEN ''
                        else t3.cust_title_thai_desc
                    END       
                    ,' '),null),
                    COALESCE(CONCAT(t3.cust_thai_name,' '),null)),
                    COALESCE(t3.cust_thai_last_name,null)))  AS customer_name,  
                    t4.bu_code as bu_code,
                    DECODE(trim(t7.BU_CODE),'NWM',t7.div_code,'SW',t7.div_code,'SCM',t7.div_code,'SBM',t7.div_code,t7.bc_code) AS  div_bc_code,
                    DECODE(trim(t7.BU_CODE),'NWM',t6.div_thai_name,'SW',t6.div_thai_name,'SCM',t6.div_thai_name,'SBM',t6.div_thai_name,t8.bc_thai_name) AS div_bc_name,
                    t4.ro_arm_ao_alloc,
                    t4.ro_arm_ao_name,
                    t4.RM_AM_ALLOC,
                    t4.RM_AM_NAME ,
                    t4.m_au_tl_name
                    FROM (SELECT * FROM colloan.t_cms_acl_stock WHERE data_date = (SELECT MAX(data_date)FROM colloan.t_cms_acl_stock))t1
                    LEFT JOIN t_cms_collateral_relation t2 ON t2.linkage_no = t1.linkage_no and t2.as_of_date = t1.data_date
                    LEFT JOIN V_SND_EM_CUST_DETAIL t3 ON t3.cust_id = t2.cust_id
                    LEFT JOIN SN_DM_PORTFOLIO_HIERARCHY t4 on t4.cust_id = t3.cust_id 
                    LEFT JOIN t_bbl_division_new t6 ON t3.div_code = t6.div_code AND t3.bu_code = t6.bu_code
                    LEFT JOIN t_cis_authorized_user t7 ON t7.allocation_code = t4.ro_arm_ao_alloc and t7.end_date is null
                    LEFT JOIN t_bbl_business_center t8 ON t7.bc_code = t8.bc_code
                    WHERE  t1.COLLATERAL_TYPE = '04' and t1.COLLATERAL_SUB_TYPE1 IN (3,5) 
                    and (t1.total_out_value is null or t1.total_out_value = 0) and t1.data_date = (SELECt MAX(t2.data_date) FROM  colloan.t_cms_acl_stock t2)  and t1.data_date = i_char_as_of_date
                    and t3.status is not null and t7.div_code  = t6.div_code
                    order by t1.DATA_DATE desc;
                    END IF;



IF i_char_as_of_date is null THEN   */
--DELETE  FROM t_rdt_exc_report_012 where data_date = i_char_as_of_date;
execute immediate ' truncate table T_RDT_EXC_REPORT_012 ';

INSERT INTO t_rdt_exc_report_012 (
    DATA_DATE,
    LINKAGE_NO,
    COLLATERAL_TYPE,
    COLLATERAL_TYPE_DESC,
    COLLATERAL_SUB_TYPE1,
    STOCK_CODE,
    APPRASIAL_VALUE,
    APPRASIAL_VALUE_CURRENCY,
    APPRASIAL_DATE,
    ISSUER_NAME,
    RM_CUST_NR,
    BOS_CUST_ID,
    CUST_NAME,
    BU_CODE,
    DIV_BC_CODE,
    DIV_BC_NAME,
    RO_ARM_AO_ALLOC_CODE,
    RO_ARM_AO_NAME,
    RM_AM_ALLOCA_CODE,
    RM_AM_NAME,
    M_TL_NAME
)
SELECT  DISTINCT
                    t1.data_date,
                    t1.linkage_no as cms_linkage_no,
                    t1.collateral_type,
                    t1.collateral_desc,
                    t1.collateral_sub_type1,
                    t1.stock_code,
                    t1.total_out_value as appraisal_value,
                    t1.currency_app_out as appraisal_value_currency,
                    t1.outsource_date as appraisal_date,
                    t1.issuer_name,
                    t3.rm_cust_nr as RM_Number,
                    t3.cust_id AS bos_cutomer_id,
                     trim(CONCAT(CONCAT(COALESCE(CONCAT(
                    CASE 
                        WHEN t3.ASSOCIATION_TYPE = 'J' THEN 'บัญชีร่วม'
                        WHEN t3.ASSOCIATION_TYPE = 'U' THEN 'กิจการร่วมค้าแบบไม่จดทะเบียน'
                        WHEN t3.ASSOCIATION_TYPE = 'C' THEN 'กิจการค้าร่วม'
                        WHEN t3.cust_title_thai_desc = 'อื่นๆ' THEN ''
                        else t3.cust_title_thai_desc
                    END       
                    ,' '),null),
                    COALESCE(CONCAT(t3.cust_thai_name,' '),null)),
                    COALESCE(t3.cust_thai_last_name,null)))  AS customer_name,  
                    t4.bu_code as bu_code,
                     DECODE(trim(t7.BU_CODE),'NWM',t7.div_code,'SW',t7.div_code,'SCM',t7.div_code,'SBM',t7.div_code,t7.bc_code) AS  div_bc_code,
                    DECODE(trim(t7.BU_CODE),'NWM',t6.div_thai_name,'SW',t6.div_thai_name,'SCM',t6.div_thai_name,'SBM',t6.div_thai_name,t8.bc_thai_name) AS div_bc_name,
                    t4.ro_arm_ao_alloc,
                    t4.ro_arm_ao_name,
                    t4.RM_AM_ALLOC,
                    t4.RM_AM_NAME ,
                    t4.m_au_tl_name
                    FROM (SELECT * FROM colloan.t_cms_acl_stock WHERE data_date = (SELECT MAX(data_date)FROM colloan.t_cms_acl_stock))t1
                    LEFT JOIN t_cms_collateral_relation t2 ON t2.linkage_no = t1.linkage_no and t2.as_of_date = t1.data_date
                    LEFT JOIN V_SND_EM_CUST_DETAIL t3 ON t3.cust_id = t2.cust_id
                    LEFT JOIN SN_DM_PORTFOLIO_HIERARCHY t4 on t4.cust_id = t3.cust_id 
                    LEFT JOIN t_bbl_division_new t6 ON t3.div_code = t6.div_code AND t3.bu_code = t6.bu_code
                    LEFT JOIN t_cis_authorized_user t7 ON t7.allocation_code = t4.ro_arm_ao_alloc and t7.end_date is null
                    LEFT JOIN t_bbl_business_center t8 ON t7.bc_code = t8.bc_code
                    WHERE  t1.COLLATERAL_TYPE = '04' and t1.COLLATERAL_SUB_TYPE1 IN (3,5) /*and t1.OUTSOURCE_DATE is null */
                    and (t1.total_out_value is null or t1.total_out_value = 0) and t1.data_date = (SELECt MAX(t2.data_date) FROM  colloan.t_cms_acl_stock t2)
                    and t3.status is not null and t7.div_code  = t6.div_code
                    order by t1.DATA_DATE desc;
    --END IF;   
    COMMIT;    
    END get_t_rdt_exc_report_012;


PROCEDURE get_t_rdt_exc_report_014  AS
    rdt_exists NUMBER;
BEGIN
    --SELECT COUNT(*)
    --INTO rdt_exists
    --FROM T_RDT_EXC_REPORT_014
    --WHERE DATA_DATE = i_char_as_of_date;

    --IF (rdt_exists > 0 ) and i_char_as_of_date is not null THEN

        --DELETE  FROM T_RDT_EXC_REPORT_014 WHERE DATA_DATE = i_char_as_of_date;
        execute immediate ' truncate table T_RDT_EXC_REPORT_014 ';
        
        insert into T_RDT_EXC_REPORT_014
        select  DISTINCT to_date(sysdate-1,'dd/mm/yyyy') as DATA_DATE,
                t1.ISIC_CODE,
                t1.sales_per_year as TOTAL_INCOME_BAHT,
                t2.income_date as DATE_OF_TOTAL_INCOME_BAHT,
                t2.domestic_income as DOMESTIC_INCOME_BAHT,
                t2.export_income as EXPORT_INCOME_BAHT,
                t1.number_of_employees as LABOR,
                t1.date_as_of_number_employees as DATE_OF_LABOR,
                'BU' as GROUP_DATA,
                t1.RM_CUST_NR,
                t1.bos_cust_id as BOS_CUST_ID,
                trim(CONCAT(CONCAT(COALESCE(CONCAT(
                        CASE 
                            WHEN t1.ASSOCIATION_TYPE = 'J' THEN 'บัญชีร่วม'
                            WHEN t1.ASSOCIATION_TYPE = 'U' THEN 'กิจการร่วมค้าแบบไม่จดทะเบียน'
                            WHEN t1.ASSOCIATION_TYPE = 'C' THEN 'กิจการค้าร่วม'
                            WHEN t1.title_thai = 'อื่นๆ' THEN ''
                            else t1.title_thai
                        END       
                        ,' '),null),
                        COALESCE(CONCAT(t1.cust_name_thai,' '),null)),
                        COALESCE(t1.cust_last_name_thai,null)))  AS customer_name,                    
                t4.bu_code as bu_code,
                DECODE(trim(t4.BU_CODE),'NWM',t4.div_code,'SW',t4.div_code,'SCM',t4.div_code,'SBM',t4.div_code,t4.bc_code) AS  div_bc_code,
                DECODE(trim(t4.BU_CODE),'NWM',t5.div_thai_name,'SW',t5.div_thai_name,'SCM',t5.div_thai_name,'SBM',t5.div_thai_name,t6.bc_thai_name) AS div_bc_name,
                t3.l1_allocation_code as ro_arm_ao_alloc,
                t4.user_thai_name as ro_arm_ao_name,
                t3.l2_allocation_code as RM_AM_ALLOC,
                t7.user_thai_name as RM_AM_NAME ,
                t8.user_thai_name as m_au_tl_name
        FROM  SND_EM_CUST_DETAIL t1,
              snd_em_cust_detail_additional t2,
              ( select * 
                from t_cis_cust_responsibility_hist 
                where (cust_id,end_date) in (select cust_id,max(end_date) from t_cis_cust_responsibility_hist group by cust_id))t3,
              (select * from t_cis_authorized_user where (allocation_code,start_date) in (select allocation_code,max(start_date) from t_cis_authorized_user group by allocation_code)) t4,
              t_bbl_division_new t5,
              t_bbl_business_center t6 ,
              (select * from t_cis_authorized_user where (allocation_code,start_date) in (select allocation_code,max(start_date) from t_cis_authorized_user group by allocation_code)) t7,
              (select * from t_cis_authorized_user where (allocation_code,start_date) in (select allocation_code,max(start_date) from t_cis_authorized_user group by allocation_code)) t8,
              (select cust_id from t_cri_cust_sum_facil
                                where credit_type_code not in ('0205','0206','0209','0212','9999')
                                and  curr_code <> '999'
                                and cust_id NOT IN (0, 9999999999)
                                and as_of_date = (select max(as_of_date) from t_cri_cons_as_of_date) ) credit
        where t1.bos_cust_id = t2.bos_cust_id
        and t1.bos_cust_id not in (select cust_id from SN_DM_PORTFOLIO_HIERARCHY)
        and t1.bos_cust_id = t3.cust_id
        AND t3.l1_allocation_code = t4.allocation_code
        and t4.div_code = t5.div_code(+)
        and t4.bc_code = t6.bc_code(+)
        AND t3.l2_allocation_code = t7.allocation_code
        AND t3.l3_allocation_code = t8.allocation_code
        and t1.bos_cust_id = credit.cust_id
        and t1.customer_status = 'E';



        insert into T_RDT_EXC_REPORT_014
        select  DISTINCT to_date(sysdate-1,'dd/mm/yyyy') as DATA_DATE,
                t1.ISIC_CODE,
                t1.sales_per_year as TOTAL_INCOME_BAHT,
                t2.income_date as DATE_OF_TOTAL_INCOME_BAHT,
                t2.domestic_income as DOMESTIC_INCOME_BAHT,
                t2.export_income as EXPORT_INCOME_BAHT,
                t1.number_of_employees as LABOR,
                t1.date_as_of_number_employees as DATE_OF_LABOR,
                'CC' as GROUP_DATA,
                t1.RM_CUST_NR,
                t1.bos_cust_id as BOS_CUST_ID,
                trim(CONCAT(CONCAT(COALESCE(CONCAT(
                        CASE 
                            WHEN t1.ASSOCIATION_TYPE = 'J' THEN 'บัญชีร่วม'
                            WHEN t1.ASSOCIATION_TYPE = 'U' THEN 'กิจการร่วมค้าแบบไม่จดทะเบียน'
                            WHEN t1.ASSOCIATION_TYPE = 'C' THEN 'กิจการค้าร่วม'
                            WHEN t1.title_thai = 'อื่นๆ' THEN ''
                            else t1.title_thai
                        END       
                        ,' '),null),
                        COALESCE(CONCAT(t1.cust_name_thai,' '),null)),
                        COALESCE(t1.cust_last_name_thai,null)))  AS customer_name,                    
                null as bu_code,
                null as div_bc_code,
                null as div_bc_name,
                null as ro_arm_ao_alloc,
                null as ro_arm_ao_name,
                null as RM_AM_ALLOC,
                null as RM_AM_NAME ,
                null as m_au_tl_name
        FROM  SND_EM_CUST_DETAIL t1,
              snd_em_cust_detail_additional t2,
              t_cri_card_corporate t3
        where t1.bos_cust_id = t2.bos_cust_id
        and t1.bos_cust_id not in (select bos_cust_id from T_RDT_EXC_REPORT_014)
        and t1.bos_cust_id not in (select cust_id from SN_DM_PORTFOLIO_HIERARCHY)
        and t1.rm_cust_nr = t3.rm_cust_nr
        and t3.as_of_date = (select max(as_of_date) from t_cri_cons_as_of_date)
        --and t1.customer_status = 'E'
        ;


        insert into T_RDT_EXC_REPORT_014
        select  DISTINCT to_date(sysdate-1,'dd/mm/yyyy') as DATA_DATE,
                t1.ISIC_CODE,
                t1.sales_per_year as TOTAL_INCOME_BAHT,
                t2.income_date as DATE_OF_TOTAL_INCOME_BAHT,
                t2.domestic_income as DOMESTIC_INCOME_BAHT,
                t2.export_income as EXPORT_INCOME_BAHT,
                t1.number_of_employees as LABOR,
                t1.date_as_of_number_employees as DATE_OF_LABOR,
                null as GROUP_DATA,
                t1.RM_CUST_NR,
                t1.bos_cust_id as BOS_CUST_ID,
                trim(CONCAT(CONCAT(COALESCE(CONCAT(
                        CASE 
                            WHEN t1.ASSOCIATION_TYPE = 'J' THEN 'บัญชีร่วม'
                            WHEN t1.ASSOCIATION_TYPE = 'U' THEN 'กิจการร่วมค้าแบบไม่จดทะเบียน'
                            WHEN t1.ASSOCIATION_TYPE = 'C' THEN 'กิจการค้าร่วม'
                            WHEN t1.title_thai = 'อื่นๆ' THEN ''
                            else t1.title_thai
                        END       
                        ,' '),null),
                        COALESCE(CONCAT(t1.cust_name_thai,' '),null)),
                        COALESCE(t1.cust_last_name_thai,null)))  AS customer_name,                    
                null as bu_code,
                null as div_bc_code,
                null as div_bc_name,
                null as ro_arm_ao_alloc,
                null as ro_arm_ao_name,
                null as RM_AM_ALLOC,
                null as RM_AM_NAME ,
                null as m_au_tl_name
        FROM  SND_EM_CUST_DETAIL t1,
              snd_em_cust_detail_additional t2,
              (select cust_id from t_cri_cust_sum_facil
                                where credit_type_code not in ('0205','0206','0209','0212','9999')
                                and  curr_code <> '999'
                                and cust_id NOT IN (0, 9999999999)
                                and as_of_date = (select max(as_of_date) from t_cri_cons_as_of_date) ) credit
        where t1.bos_cust_id = t2.bos_cust_id
        and t1.bos_cust_id not in (select bos_cust_id from T_RDT_EXC_REPORT_014)
        and t1.bos_cust_id not in (select cust_id from SN_DM_PORTFOLIO_HIERARCHY)
        and t1.rm_cust_nr not in (select rm_cust_nr from t_cri_card_corporate where as_of_date = (select max(as_of_date) from t_cri_cons_as_of_date))
        and t1.bos_cust_id = credit.cust_id
        and t1.customer_status = 'E';


    --END IF;
    COMMIT;  
END get_t_rdt_exc_report_014;       

PROCEDURE get_t_rdt_exc_report_015  AS
    rdt_exists NUMBER;
BEGIN
    --SELECT COUNT(*)
    --INTO rdt_exists
    --FROM T_RDT_EXC_REPORT_015
    --WHERE DATA_DATE = i_char_as_of_date;

    --IF (rdt_exists > 0 ) and i_char_as_of_date is not null THEN

        --DELETE  FROM T_RDT_EXC_REPORT_015 WHERE DATA_DATE = i_char_as_of_date;
        execute immediate ' truncate table T_RDT_EXC_REPORT_015 ';
            
        insert into T_RDT_EXC_REPORT_015
        select  DISTINCT to_date(sysdate-1,'dd/mm/yyyy') as DATA_DATE,
                t1.ISIC_CODE,
                t1.sales_per_year as TOTAL_INCOME_BAHT,
                t2.income_date as DATE_OF_TOTAL_INCOME_BAHT,
                t2.domestic_income as DOMESTIC_INCOME_BAHT,
                t2.export_income as EXPORT_INCOME_BAHT,
                t1.number_of_employees as LABOR,
                t1.date_as_of_number_employees as DATE_OF_LABOR,
                'BU' as GROUP_DATA,
                t1.RM_CUST_NR,
                t1.bos_cust_id as BOS_CUST_ID,
                trim(CONCAT(CONCAT(COALESCE(CONCAT(
                        CASE 
                            WHEN t1.ASSOCIATION_TYPE = 'J' THEN 'บัญชีร่วม'
                            WHEN t1.ASSOCIATION_TYPE = 'U' THEN 'กิจการร่วมค้าแบบไม่จดทะเบียน'
                            WHEN t1.ASSOCIATION_TYPE = 'C' THEN 'กิจการค้าร่วม'
                            WHEN t1.title_thai = 'อื่นๆ' THEN ''
                            else t1.title_thai
                        END       
                        ,' '),null),
                        COALESCE(CONCAT(t1.cust_name_thai,' '),null)),
                        COALESCE(t1.cust_last_name_thai,null)))  AS customer_name,                    
                t3.bu_code as bu_code,
                DECODE(trim(t4.BU_CODE),'NWM',t4.div_code,'SW',t4.div_code,'SCM',t4.div_code,'SBM',t4.div_code,t4.bc_code) AS  div_bc_code,
                DECODE(trim(t4.BU_CODE),'NWM',t5.div_thai_name,'SW',t5.div_thai_name,'SCM',t5.div_thai_name,'SBM',t5.div_thai_name,t6.bc_thai_name) AS div_bc_name,
                t3.ro_arm_ao_alloc,
                t3.ro_arm_ao_name,
                t3.RM_AM_ALLOC,
                t3.RM_AM_NAME ,
                t3.m_au_tl_name
        FROM  SND_EM_CUST_DETAIL t1,
              snd_em_cust_detail_additional t2,
              SN_DM_PORTFOLIO_HIERARCHY t3,
              t_cis_authorized_user t4,
              t_bbl_division_new t5,
              t_bbl_business_center t6 ,
              (select cust_id from t_cri_cust_sum_facil
                                where credit_type_code not in ('0205','0206','0209','0212','9999')
                                and  curr_code <> '999'
                                and cust_id NOT IN (0, 9999999999)
                                and as_of_date = (select max(as_of_date) from t_cri_cons_as_of_date) ) credit
        where t1.bos_cust_id = t2.bos_cust_id
        and t1.bos_cust_id = t3.cust_id
        AND t3.ro_arm_ao_alloc = t4.allocation_code and t4.end_date is null
        and t4.div_code = t5.div_code(+)
        and t4.bc_code = t6.bc_code(+)
        and t1.bos_cust_id  = credit.cust_id
        and t1.customer_status = 'E'
        and t1.sales_per_year is null;



        insert into T_RDT_EXC_REPORT_015
        select  DISTINCT to_date(sysdate-1,'dd/mm/yyyy') as DATA_DATE,
                t1.ISIC_CODE,
                t1.sales_per_year as TOTAL_INCOME_BAHT,
                t2.income_date as DATE_OF_TOTAL_INCOME_BAHT,
                t2.domestic_income as DOMESTIC_INCOME_BAHT,
                t2.export_income as EXPORT_INCOME_BAHT,
                t1.number_of_employees as LABOR,
                t1.date_as_of_number_employees as DATE_OF_LABOR,
                'CC' as GROUP_DATA,
                t1.RM_CUST_NR,
                t1.bos_cust_id as BOS_CUST_ID,
                trim(CONCAT(CONCAT(COALESCE(CONCAT(
                        CASE 
                            WHEN t1.ASSOCIATION_TYPE = 'J' THEN 'บัญชีร่วม'
                            WHEN t1.ASSOCIATION_TYPE = 'U' THEN 'กิจการร่วมค้าแบบไม่จดทะเบียน'
                            WHEN t1.ASSOCIATION_TYPE = 'C' THEN 'กิจการค้าร่วม'
                            WHEN t1.title_thai = 'อื่นๆ' THEN ''
                            else t1.title_thai
                        END       
                        ,' '),null),
                        COALESCE(CONCAT(t1.cust_name_thai,' '),null)),
                        COALESCE(t1.cust_last_name_thai,null)))  AS customer_name,                    
                null as bu_code,
                null as div_bc_code,
                null as div_bc_name,
                null as ro_arm_ao_alloc,
                null as ro_arm_ao_name,
                null as RM_AM_ALLOC,
                null as RM_AM_NAME ,
                null as m_au_tl_name
        FROM  SND_EM_CUST_DETAIL t1,
              snd_em_cust_detail_additional t2,
              t_cri_card_corporate t3
        where t1.bos_cust_id = t2.bos_cust_id
        and t1.bos_cust_id not in (select bos_cust_id from T_RDT_EXC_REPORT_015)
        and t1.bos_cust_id not in (select cust_id from SN_DM_PORTFOLIO_HIERARCHY)
        and t1.rm_cust_nr = t3.rm_cust_nr
        and t3.as_of_date = (select max(as_of_date) from t_cri_cons_as_of_date)
        and t1.sales_per_year is null;
    --END IF;
      COMMIT;  
END get_t_rdt_exc_report_015;   

PROCEDURE get_t_rdt_exc_report_016  AS
    rdt_exists NUMBER;
BEGIN
    --SELECT COUNT(*)
    --INTO rdt_exists
    --FROM T_RDT_EXC_REPORT_016
    --WHERE DATA_DATE = i_char_as_of_date;

    --IF (rdt_exists > 0 ) and i_char_as_of_date is not null THEN

        --DELETE  FROM T_RDT_EXC_REPORT_016 WHERE DATA_DATE = i_char_as_of_date;
        execute immediate ' truncate table T_RDT_EXC_REPORT_016 ';
        
        insert into T_RDT_EXC_REPORT_016
        select  DISTINCT to_date(sysdate-1,'dd/mm/yyyy') as DATA_DATE,
                t1.ISIC_CODE,
                t1.sales_per_year as TOTAL_INCOME_BAHT,
                t2.income_date as DATE_OF_TOTAL_INCOME_BAHT,
                t2.domestic_income as DOMESTIC_INCOME_BAHT,
                t2.export_income as EXPORT_INCOME_BAHT,
                t1.number_of_employees as LABOR,
                t1.date_as_of_number_employees as DATE_OF_LABOR,
                'BU' as GROUP_DATA,
                t1.RM_CUST_NR,
                t1.bos_cust_id as BOS_CUST_ID,
                trim(CONCAT(CONCAT(COALESCE(CONCAT(
                        CASE 
                            WHEN t1.ASSOCIATION_TYPE = 'J' THEN 'บัญชีร่วม'
                            WHEN t1.ASSOCIATION_TYPE = 'U' THEN 'กิจการร่วมค้าแบบไม่จดทะเบียน'
                            WHEN t1.ASSOCIATION_TYPE = 'C' THEN 'กิจการค้าร่วม'
                            WHEN t1.title_thai = 'อื่นๆ' THEN ''
                            else t1.title_thai
                        END       
                        ,' '),null),
                        COALESCE(CONCAT(t1.cust_name_thai,' '),null)),
                        COALESCE(t1.cust_last_name_thai,null)))  AS customer_name,                    
                t3.bu_code as bu_code,
                DECODE(trim(t4.BU_CODE),'NWM',t4.div_code,'SW',t4.div_code,'SCM',t4.div_code,'SBM',t4.div_code,t4.bc_code) AS  div_bc_code,
                DECODE(trim(t4.BU_CODE),'NWM',t5.div_thai_name,'SW',t5.div_thai_name,'SCM',t5.div_thai_name,'SBM',t5.div_thai_name,t6.bc_thai_name) AS div_bc_name,
                t3.ro_arm_ao_alloc,
                t3.ro_arm_ao_name,
                t3.RM_AM_ALLOC,
                t3.RM_AM_NAME ,
                t3.m_au_tl_name
        FROM  SND_EM_CUST_DETAIL t1,
              snd_em_cust_detail_additional t2,
              SN_DM_PORTFOLIO_HIERARCHY t3,
              t_cis_authorized_user t4,
              t_bbl_division_new t5,
              t_bbl_business_center t6,
              (select cust_id from t_cri_cust_sum_facil
                                where credit_type_code not in ('0205','0206','0209','0212','9999')
                                and  curr_code <> '999'
                                and cust_id NOT IN (0, 9999999999)
                                and as_of_date = (select max(as_of_date) from t_cri_cons_as_of_date) ) credit
        where t1.bos_cust_id = t2.bos_cust_id
        and t1.bos_cust_id = t3.cust_id
        AND t3.ro_arm_ao_alloc = t4.allocation_code and t4.end_date is null
        and t4.div_code = t5.div_code(+)
        and t4.bc_code = t6.bc_code(+)
        and t1.bos_cust_id  = credit.cust_id
        and t1.customer_status = 'E'
        and t1.number_of_employees is null;



        insert into T_RDT_EXC_REPORT_016
        select  DISTINCT to_date(sysdate-1,'dd/mm/yyyy') as DATA_DATE,
                t1.ISIC_CODE,
                t1.sales_per_year as TOTAL_INCOME_BAHT,
                t2.income_date as DATE_OF_TOTAL_INCOME_BAHT,
                t2.domestic_income as DOMESTIC_INCOME_BAHT,
                t2.export_income as EXPORT_INCOME_BAHT,
                t1.number_of_employees as LABOR,
                t1.date_as_of_number_employees as DATE_OF_LABOR,
                'CC' as GROUP_DATA,
                t1.RM_CUST_NR,
                t1.bos_cust_id as BOS_CUST_ID,
                trim(CONCAT(CONCAT(COALESCE(CONCAT(
                        CASE 
                            WHEN t1.ASSOCIATION_TYPE = 'J' THEN 'บัญชีร่วม'
                            WHEN t1.ASSOCIATION_TYPE = 'U' THEN 'กิจการร่วมค้าแบบไม่จดทะเบียน'
                            WHEN t1.ASSOCIATION_TYPE = 'C' THEN 'กิจการค้าร่วม'
                            WHEN t1.title_thai = 'อื่นๆ' THEN ''
                            else t1.title_thai
                        END       
                        ,' '),null),
                        COALESCE(CONCAT(t1.cust_name_thai,' '),null)),
                        COALESCE(t1.cust_last_name_thai,null)))  AS customer_name,                    
                null as bu_code,
                null as div_bc_code,
                null as div_bc_name,
                null as ro_arm_ao_alloc,
                null as ro_arm_ao_name,
                null as RM_AM_ALLOC,
                null as RM_AM_NAME ,
                null as m_au_tl_name
        FROM  SND_EM_CUST_DETAIL t1,
              snd_em_cust_detail_additional t2,
              t_cri_card_corporate t3
        where t1.bos_cust_id = t2.bos_cust_id
        and t1.bos_cust_id not in (select bos_cust_id from T_RDT_EXC_REPORT_016)
        and t1.bos_cust_id not in (select cust_id from SN_DM_PORTFOLIO_HIERARCHY)
        and t1.rm_cust_nr = t3.rm_cust_nr
        and t3.as_of_date = (select max(as_of_date) from t_cri_cons_as_of_date)
        and t1.number_of_employees is null;
      COMMIT;  
    --END IF;
END get_t_rdt_exc_report_016;   


PROCEDURE get_t_rdt_exc_report_hist  AS
BEGIN
    
    delete from T_RDT_EXC_REPORT_003_HIST where data_date = (select DISTINCT data_date from T_RDT_EXC_REPORT_003);
    delete from T_RDT_EXC_REPORT_022_HIST where data_date = (select DISTINCT data_date from T_RDT_EXC_REPORT_022);
    delete from T_RDT_EXC_REPORT_011_HIST where data_date = (select DISTINCT data_date from T_RDT_EXC_REPORT_011);
    delete from T_RDT_EXC_REPORT_012_HIST where data_date = (select DISTINCT data_date from T_RDT_EXC_REPORT_012);
    delete from T_RDT_EXC_REPORT_014_HIST where data_date = (select DISTINCT data_date from T_RDT_EXC_REPORT_014);
    delete from T_RDT_EXC_REPORT_015_HIST where data_date = (select DISTINCT data_date from T_RDT_EXC_REPORT_015);
    delete from T_RDT_EXC_REPORT_016_HIST where data_date = (select DISTINCT data_date from T_RDT_EXC_REPORT_016);

    insert into T_RDT_EXC_REPORT_003_HIST select * from T_RDT_EXC_REPORT_003;
    insert into T_RDT_EXC_REPORT_022_HIST select * from T_RDT_EXC_REPORT_022;
    insert into T_RDT_EXC_REPORT_011_HIST select * from T_RDT_EXC_REPORT_011;
    insert into T_RDT_EXC_REPORT_012_HIST select * from T_RDT_EXC_REPORT_012;
    insert into T_RDT_EXC_REPORT_014_HIST select * from T_RDT_EXC_REPORT_014;
    insert into T_RDT_EXC_REPORT_015_HIST select * from T_RDT_EXC_REPORT_015;
    insert into T_RDT_EXC_REPORT_016_HIST select * from T_RDT_EXC_REPORT_016;
    commit;
    
END;



END CRM_RDT_EXCEPTION_PKG;