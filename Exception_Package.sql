create or replace PACKAGE CRM_RDT_EXCEPTION_PKG

IS
  /*---Share---*/
PROCEDURE get_t_rdt_exc_report_022;
PROCEDURE get_t_rdt_exc_report_003;
PROCEDURE get_t_rdt_exc_report_011;
PROCEDURE get_t_rdt_exc_report_012;
PROCEDURE get_t_rdt_exc_report_014;
PROCEDURE get_t_rdt_exc_report_015;
PROCEDURE get_t_rdt_exc_report_016;
   /*PROCEDURE get_rdt_exception_report_data( i_char_as_of_date IN VARCHAR2 Default NULL );*/

PROCEDURE get_t_rdt_exc_report_hist;

END CRM_RDT_EXCEPTION_PKG;