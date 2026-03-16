REPORT ztable_impex.

TYPE-POOLS vrm.

TABLES dd02l.

CONSTANTS:
  gc_xchk_disabled TYPE c LENGTH 1 VALUE 'D',
  gc_xchk_warning  TYPE c LENGTH 1 VALUE 'W',
  gc_xchk_error    TYPE c LENGTH 1 VALUE 'E'.

TYPES:
  ty_xchk_mode TYPE c LENGTH 1,
  BEGIN OF ty_log,
    msgty TYPE symsgty,
    line  TYPE i,
    text  TYPE string,
  END OF ty_log,
  tt_log TYPE STANDARD TABLE OF ty_log WITH EMPTY KEY.

TYPES:
  BEGIN OF ty_header_map,
    excel_col   TYPE string,
    fieldname   TYPE fieldname,
    column_text TYPE string,
  END OF ty_header_map,
  tt_header_map TYPE STANDARD TABLE OF ty_header_map WITH EMPTY KEY.

TYPES:
  BEGIN OF ty_field_meta,
    fieldname   TYPE fieldname,
    keyflag     TYPE xfeld,
    position    TYPE i,
    leng        TYPE i,
    convexit    TYPE convexit,
    checktable  TYPE checktable,
    datatype    TYPE datatype_d,
  END OF ty_field_meta,
  tt_field_meta TYPE STANDARD TABLE OF ty_field_meta WITH EMPTY KEY.

PARAMETERS:
  p_file  TYPE string LOWER CASE,
  p_tab   TYPE tabname OBLIGATORY.

SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-t01.
PARAMETERS:
  p_imp RADIOBUTTON GROUP act DEFAULT 'X' USER-COMMAND act,
  p_exp RADIOBUTTON GROUP act.
PARAMETERS:
  p_dry  AS CHECKBOX DEFAULT 'X' MODIF ID imp,
  p_conv AS CHECKBOX DEFAULT 'X',
  p_ovr  AS CHECKBOX MODIF ID imp,
  p_xchk TYPE c LENGTH 1 AS LISTBOX VISIBLE LENGTH 10 DEFAULT gc_xchk_error MODIF ID imp.
SELECTION-SCREEN END OF BLOCK b1.

CLASS lcl_app DEFINITION FINAL.
  PUBLIC SECTION.
    CLASS-METHODS initialization.
    CLASS-METHODS adjust_screen.
    CLASS-METHODS run.
    CLASS-METHODS value_request_file.
  PRIVATE SECTION.
    CLASS-DATA gt_log TYPE tt_log.
    CLASS-METHODS set_xchk_values.
    CLASS-METHODS add_log
      IMPORTING
        iv_msgty TYPE symsgty
        iv_text  TYPE string
        iv_line  TYPE i DEFAULT 0.
    CLASS-METHODS has_errors
      RETURNING VALUE(rv_has_errors) TYPE abap_bool.
    CLASS-METHODS show_log
      RAISING
        cx_salv_msg.
    CLASS-METHODS validate_table
      EXPORTING
        ev_is_customizing TYPE abap_bool
        ev_client_field   TYPE fieldname
        et_fields         TYPE tt_field_meta.
    CLASS-METHODS load_field_metadata
      IMPORTING
        iv_tabname       TYPE tabname
      RETURNING
        VALUE(rt_fields) TYPE tt_field_meta.
    CLASS-METHODS build_dynamic_table
      IMPORTING
        iv_tabname   TYPE tabname
      EXPORTING
        er_table     TYPE REF TO data
        er_line      TYPE REF TO data.
    CLASS-METHODS export_table
      RAISING
        zcx_excel.
    CLASS-METHODS import_table
      RAISING
        zcx_excel
        cx_salv_msg.
    CLASS-METHODS build_header_map
      IMPORTING
        io_worksheet          TYPE REF TO object
        it_fields             TYPE tt_field_meta
      EXPORTING
        et_header_map         TYPE tt_header_map
        ev_highest_row        TYPE i
        ev_key_missing        TYPE abap_bool
      RAISING
        zcx_excel.
    CLASS-METHODS column_index_to_alpha
      IMPORTING
        iv_index        TYPE i
      RETURNING
        VALUE(rv_alpha) TYPE string.
    CLASS-METHODS column_alpha_to_index
      IMPORTING
        iv_alpha        TYPE string
      RETURNING
        VALUE(rv_index) TYPE i.
    CLASS-METHODS apply_conversion_exit
      IMPORTING
        iv_value        TYPE string
        iv_convexit     TYPE convexit
        iv_direction    TYPE string
      RETURNING
        VALUE(rv_value) TYPE string.
    CLASS-METHODS component_to_string
      IMPORTING
        is_field       TYPE ty_field_meta
      CHANGING
        cv_value       TYPE any
      RETURNING
        VALUE(rv_text) TYPE string.
    CLASS-METHODS is_row_initial
      IMPORTING
        is_row               TYPE any
        it_header_map        TYPE tt_header_map
      RETURNING
        VALUE(rv_is_initial) TYPE abap_bool.
    CLASS-METHODS build_where_from_row
      IMPORTING
        is_row          TYPE any
        it_fields       TYPE tt_field_meta
      RETURNING
        VALUE(rv_where) TYPE string.
    CLASS-METHODS sql_quote
      IMPORTING
        iv_value        TYPE string
      RETURNING
        VALUE(rv_value) TYPE string.
    CLASS-METHODS check_check_tables
      IMPORTING
        is_row        TYPE any
        iv_excel_line TYPE i
        it_fields     TYPE tt_field_meta
        iv_mode       TYPE ty_xchk_mode.
    CLASS-METHODS register_transport_keys
      IMPORTING
        iv_tabname      TYPE tabname
        iv_transport    TYPE trkorr
        iv_client_field TYPE fieldname
        it_fields       TYPE tt_field_meta
        ir_rows         TYPE REF TO data.
    CLASS-METHODS file_open_dialog
      RETURNING VALUE(rv_path) TYPE string.
    CLASS-METHODS file_save_dialog
      IMPORTING
        iv_default_name TYPE string
      RETURNING VALUE(rv_path) TYPE string.
ENDCLASS.

AT SELECTION-SCREEN OUTPUT.
  lcl_app=>adjust_screen( ).

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  lcl_app=>value_request_file( ).

INITIALIZATION.
  lcl_app=>initialization( ).

START-OF-SELECTION.
  lcl_app=>run( ).

CLASS lcl_app IMPLEMENTATION.
  METHOD initialization.
    set_xchk_values( ).
  ENDMETHOD.

  METHOD adjust_screen.
    set_xchk_values( ).
    LOOP AT SCREEN.
      IF screen-group1 = 'IMP'.
        screen-active = COND #( WHEN p_imp = abap_true THEN 1 ELSE 0 ).
        MODIFY SCREEN.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.

  METHOD value_request_file.
    p_file = file_open_dialog( ).
  ENDMETHOD.

  METHOD run.
    CLEAR gt_log.

    TRY.
        IF p_exp = abap_true.
          export_table( ).
        ELSE.
          import_table( ).
        ENDIF.
      CATCH zcx_excel cx_salv_msg INTO DATA(lx_error).
        MESSAGE lx_error->get_text( ) TYPE 'S' DISPLAY LIKE 'E'.
    ENDTRY.
  ENDMETHOD.

  METHOD set_xchk_values.
    DATA lt_values TYPE vrm_values.

    lt_values = VALUE #(
      ( key = gc_xchk_disabled text = 'disabled' )
      ( key = gc_xchk_warning  text = 'warning' )
      ( key = gc_xchk_error    text = 'error' ) ).

    CALL FUNCTION 'VRM_SET_VALUES'
      EXPORTING
        id     = 'P_XCHK'
        values = lt_values.
  ENDMETHOD.

  METHOD add_log.
    APPEND VALUE #( msgty = iv_msgty line = iv_line text = iv_text ) TO gt_log.
  ENDMETHOD.

  METHOD has_errors.
    rv_has_errors = xsdbool( line_exists( gt_log[ msgty = 'E' ] ) ).
  ENDMETHOD.

  METHOD show_log.
    DATA lo_alv TYPE REF TO cl_salv_table.

    IF gt_log IS INITIAL.
      RETURN.
    ENDIF.

    cl_salv_table=>factory(
      IMPORTING
        r_salv_table = lo_alv
      CHANGING
        t_table      = gt_log ).

    lo_alv->get_columns( )->set_optimize( abap_true ).
    lo_alv->set_screen_popup(
      start_column = 5
      end_column   = 160
      start_line   = 2
      end_line     = 25 ).
    lo_alv->display( ).
  ENDMETHOD.

  METHOD validate_table.
    DATA ls_dd02l TYPE dd02l.

    SELECT SINGLE * FROM dd02l
      INTO ls_dd02l
      WHERE tabname  = p_tab
        AND as4local = 'A'.

    IF sy-subrc <> 0.
      MESSAGE e398(00) WITH |Table { p_tab } not found|.
    ENDIF.

    IF ls_dd02l-tabclass <> 'TRANSP'.
      MESSAGE e398(00) WITH |Table { p_tab } is not a transparent table|.
    ENDIF.

    et_fields = load_field_metadata( p_tab ).
    IF et_fields IS INITIAL.
      MESSAGE e398(00) WITH |Could not read DDIC metadata for { p_tab }|.
    ENDIF.

    ev_client_field = VALUE #( ).
    READ TABLE et_fields WITH KEY fieldname = 'MANDT' INTO DATA(ls_client_field).
    IF sy-subrc = 0.
      ev_client_field = ls_client_field-fieldname.
    ENDIF.

    ev_is_customizing = xsdbool( ls_dd02l-contflag = 'C' OR ls_dd02l-contflag = 'G' ).
  ENDMETHOD.

  METHOD load_field_metadata.
    DATA lt_dfies TYPE STANDARD TABLE OF dfies.

    CALL FUNCTION 'DDIF_FIELDINFO_GET'
      EXPORTING
        tabname        = iv_tabname
        all_types      = abap_true
      TABLES
        dfies_tab      = lt_dfies
      EXCEPTIONS
        not_found      = 1
        internal_error = 2
        OTHERS         = 3.

    IF sy-subrc <> 0.
      RETURN.
    ENDIF.

    LOOP AT lt_dfies INTO DATA(ls_dfies).
      IF ls_dfies-fieldname CP '.INCLU*' OR ls_dfies-fieldname IS INITIAL.
        CONTINUE.
      ENDIF.

      APPEND VALUE #(
        fieldname  = ls_dfies-fieldname
        keyflag    = ls_dfies-keyflag
        position   = lines( rt_fields ) + 1
        leng       = ls_dfies-leng
        convexit   = ls_dfies-convexit
        checktable = ls_dfies-checktable
        datatype   = ls_dfies-datatype ) TO rt_fields.
    ENDLOOP.
  ENDMETHOD.

  METHOD build_dynamic_table.
    CREATE DATA er_table TYPE TABLE OF (iv_tabname).
    CREATE DATA er_line TYPE (iv_tabname).
  ENDMETHOD.

  METHOD export_table.
    DATA lt_fields TYPE tt_field_meta.
    DATA lv_customizing TYPE abap_bool.
    DATA lv_client_field TYPE fieldname.
    DATA lr_table TYPE REF TO data.
    DATA lr_line TYPE REF TO data.
    DATA lo_excel TYPE REF TO zcl_excel.
    DATA lo_writer TYPE REF TO zif_excel_writer.
    DATA lo_sheet TYPE REF TO zcl_excel_worksheet.
    DATA lt_binary TYPE solix_tab.
    DATA lv_xstring TYPE xstring.
    DATA lv_size TYPE i.
    DATA lv_path TYPE string.

    FIELD-SYMBOLS:
      <lt_table> TYPE STANDARD TABLE,
      <ls_row>   TYPE any,
      <lv_comp>  TYPE any.

    validate_table(
      IMPORTING
        ev_is_customizing = lv_customizing
        ev_client_field   = lv_client_field
        et_fields         = lt_fields ).

    build_dynamic_table(
      EXPORTING
        iv_tabname = p_tab
      IMPORTING
        er_table   = lr_table
        er_line    = lr_line ).

    ASSIGN lr_table->* TO <lt_table>.

    SELECT * FROM (p_tab)
      INTO TABLE <lt_table>.

    CREATE OBJECT lo_excel.
    lo_sheet = lo_excel->get_active_worksheet( ).
    lo_sheet->set_title( ip_title = CONV zexcel_sheet_title( p_tab ) ).

    LOOP AT lt_fields INTO DATA(ls_field).
      lo_sheet->set_cell(
        ip_column = column_index_to_alpha( ls_field-position )
        ip_row    = 1
        ip_value  = ls_field-fieldname ).
    ENDLOOP.

    DATA(lv_excel_row) = 1.
    LOOP AT <lt_table> ASSIGNING <ls_row>.
      lv_excel_row = lv_excel_row + 1.
      LOOP AT lt_fields INTO ls_field.
        ASSIGN COMPONENT ls_field-fieldname OF STRUCTURE <ls_row> TO <lv_comp>.
        IF sy-subrc <> 0.
          CONTINUE.
        ENDIF.

        lo_sheet->set_cell(
          ip_column = column_index_to_alpha( ls_field-position )
          ip_row    = lv_excel_row
          ip_value  = component_to_string( EXPORTING is_field = ls_field CHANGING cv_value = <lv_comp> ) ).
      ENDLOOP.
    ENDLOOP.

    lv_path = file_save_dialog( |{ p_tab }.xlsx| ).
    IF lv_path IS INITIAL.
      RETURN.
    ENDIF.

    CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
    lv_xstring = lo_writer->write_file( lo_excel ).

    CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
      EXPORTING
        buffer        = lv_xstring
      IMPORTING
        output_length = lv_size
      TABLES
        binary_tab    = lt_binary.

    cl_gui_frontend_services=>gui_download(
      EXPORTING
        bin_filesize = lv_size
        filename     = lv_path
        filetype     = 'BIN'
      CHANGING
        data_tab     = lt_binary ).

    MESSAGE s398(00) WITH |Exported { lines( <lt_table> ) } row(s) from { p_tab } to { lv_path }|.
  ENDMETHOD.

  METHOD import_table.
    DATA lt_fields TYPE tt_field_meta.
    DATA lv_is_customizing TYPE abap_bool.
    DATA lv_client_field TYPE fieldname.
    DATA lr_table TYPE REF TO data.
    DATA lr_line TYPE REF TO data.
    DATA lr_existing TYPE REF TO data.
    DATA lo_reader TYPE REF TO zif_excel_reader.
    DATA lo_excel TYPE REF TO zcl_excel.
    DATA lo_sheet TYPE REF TO zcl_excel_worksheet.
    DATA lt_header_map TYPE tt_header_map.
    DATA lv_highest_row TYPE i.
    DATA lv_key_missing TYPE abap_bool.
    DATA lv_inserted TYPE i.
    DATA lv_updated TYPE i.
    DATA lv_transport TYPE trkorr.

    FIELD-SYMBOLS:
      <lt_upload>   TYPE STANDARD TABLE,
      <ls_upload>   TYPE any,
      <ls_existing> TYPE any,
      <lv_comp>     TYPE any.

    validate_table(
      IMPORTING
        ev_is_customizing = lv_is_customizing
        ev_client_field   = lv_client_field
        et_fields         = lt_fields ).

    IF p_file IS INITIAL.
      MESSAGE e398(00) WITH 'Select an Excel file first'.
    ENDIF.

    add_log( iv_msgty = 'I' iv_text = |Action: Import| ).
    add_log( iv_msgty = 'I' iv_text = |File: { p_file }| ).
    add_log( iv_msgty = 'I' iv_text = |Table: { p_tab }| ).
    add_log( iv_msgty = 'I' iv_text = |Dry run: { COND string( WHEN p_dry = abap_true THEN 'X' ELSE '-' ) }| ).
    add_log( iv_msgty = 'I' iv_text = |Conversion exits: { COND string( WHEN p_conv = abap_true THEN 'X' ELSE '-' ) }| ).
    add_log( iv_msgty = 'I' iv_text = |Overwrite existing entries: { COND string( WHEN p_ovr = abap_true THEN 'X' ELSE '-' ) }| ).
    add_log( iv_msgty = 'I' iv_text = |Cross-check mode: { p_xchk }| ).

    build_dynamic_table(
      EXPORTING
        iv_tabname = p_tab
      IMPORTING
        er_table   = lr_table
        er_line    = lr_line ).

    CREATE DATA lr_existing TYPE (p_tab).

    ASSIGN lr_table->* TO <lt_upload>.
    ASSIGN lr_existing->* TO <ls_existing>.

    CREATE OBJECT lo_reader TYPE zcl_excel_reader_2007.
    lo_excel = lo_reader->load_file( i_filename = p_file ).
    lo_sheet = lo_excel->get_worksheet_by_index( iv_index = 1 ).

    build_header_map(
      EXPORTING
        io_worksheet   = lo_sheet
        it_fields      = lt_fields
      IMPORTING
        et_header_map  = lt_header_map
        ev_highest_row = lv_highest_row
        ev_key_missing = lv_key_missing ).

    IF lv_highest_row <= 1.
      add_log( iv_msgty = 'W' iv_text = 'The Excel file contains no data rows' ).
    ENDIF.

    IF lt_header_map IS INITIAL.
      add_log( iv_msgty = 'E' iv_text = 'Line 1 does not contain technical field names of the target table' ).
    ENDIF.

    IF lv_key_missing = abap_true.
      add_log( iv_msgty = 'E' iv_text = 'The Excel header does not contain the full table key' ).
    ENDIF.

    DO lv_highest_row - 1 TIMES.
      DATA(lv_excel_line) = sy-index + 1.

      ASSIGN lr_line->* TO <ls_upload>.
      CLEAR <ls_upload>.

      LOOP AT lt_header_map INTO DATA(ls_header_map).
        DATA(lv_cell_value) = VALUE string( ).

        lo_sheet->get_cell(
          EXPORTING
            ip_column = ls_header_map-excel_col
            ip_row    = lv_excel_line
          IMPORTING
            ep_value  = lv_cell_value ).

        READ TABLE lt_fields WITH KEY fieldname = ls_header_map-fieldname INTO DATA(ls_field).
        IF sy-subrc <> 0.
          CONTINUE.
        ENDIF.

        ASSIGN COMPONENT ls_header_map-fieldname OF STRUCTURE <ls_upload> TO <lv_comp>.
        IF sy-subrc <> 0.
          CONTINUE.
        ENDIF.

        IF p_conv = abap_true AND ls_field-convexit IS NOT INITIAL AND lv_cell_value IS NOT INITIAL.
          lv_cell_value = apply_conversion_exit(
            iv_value     = lv_cell_value
            iv_convexit  = ls_field-convexit
            iv_direction = 'INPUT' ).
        ENDIF.

        TRY.
            <lv_comp> = lv_cell_value.
          CATCH cx_sy_conversion_error INTO DATA(lx_conv).
            add_log(
              iv_msgty = 'E'
              iv_line  = lv_excel_line
              iv_text  = |Field { ls_header_map-fieldname }: { lx_conv->get_text( ) }| ).
        ENDTRY.
      ENDLOOP.

      IF lv_client_field IS NOT INITIAL.
        ASSIGN COMPONENT lv_client_field OF STRUCTURE <ls_upload> TO <lv_comp>.
        IF sy-subrc = 0 AND <lv_comp> IS INITIAL.
          <lv_comp> = sy-mandt.
        ENDIF.
      ENDIF.

      IF is_row_initial( is_row = <ls_upload> it_header_map = lt_header_map ) = abap_true.
        CONTINUE.
      ENDIF.

      check_check_tables(
        EXPORTING
          is_row        = <ls_upload>
          iv_excel_line = lv_excel_line
          it_fields     = lt_fields
          iv_mode       = p_xchk ).

      DATA(lv_where) = build_where_from_row(
        is_row    = <ls_upload>
        it_fields = lt_fields ).

      IF lv_where IS INITIAL.
        add_log( iv_msgty = 'E' iv_line = lv_excel_line iv_text = 'Could not build a key predicate for the row' ).
      ELSE.
        CLEAR <ls_existing>.
        SELECT SINGLE * FROM (p_tab)
          INTO <ls_existing>
          WHERE (lv_where).

        IF sy-subrc = 0.
          IF p_ovr = abap_true.
            lv_updated = lv_updated + 1.
            add_log( iv_msgty = 'W' iv_line = lv_excel_line iv_text = |Existing entry will be overwritten ({ lv_where })| ).
          ELSE.
            add_log( iv_msgty = 'E' iv_line = lv_excel_line iv_text = |Existing entry found and overwrite is disabled ({ lv_where })| ).
          ENDIF.
        ELSE.
          lv_inserted = lv_inserted + 1.
        ENDIF.
      ENDIF.

      APPEND <ls_upload> TO <lt_upload>.
    ENDDO.

    add_log( iv_msgty = 'I' iv_text = |Excel data rows: { lines( <lt_upload> ) }| ).

    show_log( ).

    IF has_errors( ) = abap_true OR p_dry = abap_true.
      RETURN.
    ENDIF.

    IF lv_is_customizing = abap_true.
      CALL FUNCTION 'TRINT_ORDER_CHOICE'
        EXPORTING
          wi_order_type          = 'K'
          wi_task_type           = 'S'
          wi_category            = 'CUST'
        IMPORTING
          we_order               = lv_transport
        EXCEPTIONS
          no_correction_selected = 1
          display_mode           = 2
          object_append_error    = 3
          recursive_call         = 4
          wrong_order_type       = 5
          OTHERS                 = 6.

      IF sy-subrc = 1.
        MESSAGE s398(00) WITH 'No transport selected'.
        RETURN.
      ELSEIF sy-subrc <> 0.
        MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
          WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
      ENDIF.

      register_transport_keys(
        EXPORTING
          iv_tabname      = p_tab
          iv_transport    = lv_transport
          iv_client_field = lv_client_field
          it_fields       = lt_fields
          ir_rows         = lr_table ).
    ENDIF.

    MODIFY (p_tab) FROM TABLE <lt_upload>.
    IF sy-subrc <> 0.
      MESSAGE e398(00) WITH |Database update for { p_tab } failed|.
    ENDIF.

    COMMIT WORK.
    MESSAGE s398(00) WITH |Inserted { lv_inserted } row(s), updated { lv_updated } row(s)|.
  ENDMETHOD.

  METHOD build_header_map.
    DATA lo_worksheet TYPE REF TO zcl_excel_worksheet.
    DATA lt_headers TYPE STANDARD TABLE OF string WITH EMPTY KEY.
    DATA lv_highest_col TYPE string.

    lo_worksheet ?= io_worksheet.
    ev_key_missing = abap_false.
    lv_highest_col = lo_worksheet->get_highest_column( ).
    ev_highest_row = lo_worksheet->get_highest_row( ).

    DO column_alpha_to_index( lv_highest_col ) TIMES.
      DATA(lv_column_alpha) = column_index_to_alpha( sy-index ).
      DATA(lv_header_value) = VALUE string( ).

      lo_worksheet->get_cell(
        EXPORTING
          ip_column = lv_column_alpha
          ip_row    = 1
        IMPORTING
          ep_value  = lv_header_value ).

      CONDENSE lv_header_value.
      APPEND lv_header_value TO lt_headers.

      IF lv_header_value IS INITIAL.
        add_log( iv_msgty = 'W' iv_text = |Column { lv_column_alpha } has an empty header| ).
        CONTINUE.
      ENDIF.

      READ TABLE it_fields WITH KEY fieldname = lv_header_value INTO DATA(ls_field).
      IF sy-subrc <> 0.
        add_log( iv_msgty = 'W' iv_text = |Excel column { lv_column_alpha } does not exist in table { p_tab }: { lv_header_value }| ).
        CONTINUE.
      ENDIF.

      APPEND VALUE #(
        excel_col   = lv_column_alpha
        fieldname   = ls_field-fieldname
        column_text = lv_header_value ) TO et_header_map.
    ENDDO.

    LOOP AT it_fields INTO ls_field.
      IF line_exists( et_header_map[ fieldname = ls_field-fieldname ] ).
        CONTINUE.
      ENDIF.

      add_log( iv_msgty = 'W' iv_text = |Table field { ls_field-fieldname } is missing in the Excel header| ).
      IF ls_field-keyflag = abap_true.
        ev_key_missing = abap_true.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.

  METHOD column_index_to_alpha.
    CONSTANTS lc_letters TYPE string VALUE 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.
    DATA lv_index TYPE i.
    DATA lv_rest TYPE i.

    lv_index = iv_index.
    WHILE lv_index > 0.
      lv_rest = ( lv_index - 1 ) MOD 26.
      rv_alpha = lc_letters+lv_rest(1) && rv_alpha.
      lv_index = ( lv_index - 1 ) DIV 26.
    ENDWHILE.
  ENDMETHOD.

  METHOD column_alpha_to_index.
    CONSTANTS lc_letters TYPE string VALUE 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.
    DATA lv_alpha TYPE string.
    DATA lv_char TYPE c LENGTH 1.
    DATA lv_offset TYPE i.
    DATA lv_pos TYPE i.

    lv_alpha = to_upper( iv_alpha ).
    DO strlen( lv_alpha ) TIMES.
      lv_offset = sy-index - 1.
      lv_char = lv_alpha+lv_offset(1).
      FIND lv_char IN lc_letters MATCH OFFSET lv_pos.
      rv_index = rv_index * 26 + lv_pos + 1.
    ENDDO.
  ENDMETHOD.

  METHOD apply_conversion_exit.
    DATA lv_fm TYPE rs38l_fnam.

    rv_value = iv_value.
    IF iv_convexit IS INITIAL OR iv_value IS INITIAL.
      RETURN.
    ENDIF.

    lv_fm = |CONVERSION_EXIT_{ iv_convexit }_{ iv_direction }|.

    CALL FUNCTION lv_fm
      EXPORTING
        input  = iv_value
      IMPORTING
        output = rv_value
      EXCEPTIONS
        OTHERS = 1.
  ENDMETHOD.

  METHOD component_to_string.
    DATA lv_text TYPE string.

    lv_text = |{ cv_value }|.
    IF p_conv = abap_true AND is_field-convexit IS NOT INITIAL AND lv_text IS NOT INITIAL.
      lv_text = apply_conversion_exit(
        iv_value     = lv_text
        iv_convexit  = is_field-convexit
        iv_direction = 'OUTPUT' ).
    ENDIF.

    rv_text = lv_text.
  ENDMETHOD.

  METHOD is_row_initial.
    FIELD-SYMBOLS <lv_comp> TYPE any.

    rv_is_initial = abap_true.
    LOOP AT it_header_map INTO DATA(ls_header_map).
      ASSIGN COMPONENT ls_header_map-fieldname OF STRUCTURE is_row TO <lv_comp>.
      IF sy-subrc = 0 AND <lv_comp> IS NOT INITIAL.
        rv_is_initial = abap_false.
        EXIT.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.

  METHOD build_where_from_row.
    FIELD-SYMBOLS <lv_comp> TYPE any.

    LOOP AT it_fields INTO DATA(ls_field) WHERE keyflag = abap_true.
      ASSIGN COMPONENT ls_field-fieldname OF STRUCTURE is_row TO <lv_comp>.
      IF sy-subrc <> 0.
        CONTINUE.
      ENDIF.

      DATA(lv_value) = |{ <lv_comp> }|.
      IF rv_where IS NOT INITIAL.
        rv_where = rv_where && ` AND `.
      ENDIF.
      rv_where = rv_where && |{ ls_field-fieldname } = { sql_quote( lv_value ) }|.
    ENDLOOP.
  ENDMETHOD.

  METHOD sql_quote.
    rv_value = iv_value.
    REPLACE ALL OCCURRENCES OF '''' IN rv_value WITH ''''''.
    rv_value = |'{ rv_value }'|.
  ENDMETHOD.

  METHOD check_check_tables.
    DATA lt_check_fields TYPE tt_field_meta.
    DATA lv_checktable TYPE tabname.
    DATA lv_count TYPE i.

    FIELD-SYMBOLS <lv_comp> TYPE any.

    IF iv_mode = gc_xchk_disabled.
      RETURN.
    ENDIF.

    LOOP AT it_fields INTO DATA(ls_field) WHERE checktable IS NOT INITIAL.
      ASSIGN COMPONENT ls_field-fieldname OF STRUCTURE is_row TO <lv_comp>.
      IF sy-subrc <> 0 OR <lv_comp> IS INITIAL.
        CONTINUE.
      ENDIF.

      lt_check_fields = load_field_metadata( ls_field-checktable ).
      DELETE lt_check_fields WHERE keyflag <> abap_true OR fieldname = 'MANDT'.

      IF lt_check_fields IS INITIAL.
        CONTINUE.
      ENDIF.

      READ TABLE lt_check_fields WITH KEY fieldname = ls_field-fieldname TRANSPORTING NO FIELDS.
      IF sy-subrc <> 0.
        add_log(
          iv_msgty = COND #( WHEN iv_mode = gc_xchk_error THEN 'E' ELSE 'W' )
          iv_line  = iv_excel_line
          iv_text  = |Cross-check skipped for field { ls_field-fieldname }: check table { ls_field-checktable } uses a different key field name| ).
        CONTINUE.
      ENDIF.

      DATA(lv_where) = |{ ls_field-fieldname } = { sql_quote( |{ <lv_comp> }| ) }|.
      lv_checktable = ls_field-checktable.
      SELECT COUNT( * ) FROM (lv_checktable)
        INTO lv_count
        WHERE (lv_where).

      IF lv_count = 0.
        add_log(
          iv_msgty = COND #( WHEN iv_mode = gc_xchk_error THEN 'E' ELSE 'W' )
          iv_line  = iv_excel_line
          iv_text  = |No check table entry in { ls_field-checktable } for { ls_field-fieldname } = { <lv_comp> }| ).
      ENDIF.
    ENDLOOP.
  ENDMETHOD.

  METHOD register_transport_keys.
    DATA lt_ko200 TYPE STANDARD TABLE OF ko200.
    DATA lt_e071k TYPE STANDARD TABLE OF e071k.
    DATA ls_ko200 TYPE ko200.
    DATA ls_e071k TYPE e071k.
    DATA lv_offset TYPE i.

    FIELD-SYMBOLS:
      <lt_rows>     TYPE STANDARD TABLE,
      <ls_row>      TYPE any,
      <lv_comp>     TYPE any,
      <lv_tab_part> TYPE any.

    ASSIGN ir_rows->* TO <lt_rows>.

    ls_ko200-pgmid = 'R3TR'.
    ls_ko200-object = 'TABU'.
    ls_ko200-objfunc = 'K'.
    ls_ko200-obj_name = iv_tabname.
    APPEND ls_ko200 TO lt_ko200.

    LOOP AT <lt_rows> ASSIGNING <ls_row>.
      CLEAR ls_e071k.
      ls_e071k-pgmid = 'R3TR'.
      ls_e071k-object = 'TABU'.
      ls_e071k-mastertype = 'TABU'.
      ls_e071k-objname = iv_tabname.
      ls_e071k-mastername = iv_tabname.

      lv_offset = 0.
      LOOP AT it_fields INTO DATA(ls_field) WHERE keyflag = abap_true.
        ASSIGN COMPONENT ls_field-fieldname OF STRUCTURE <ls_row> TO <lv_comp>.
        IF sy-subrc <> 0.
          CONTINUE.
        ENDIF.

        ASSIGN ls_e071k-tabkey+lv_offset(ls_field-leng) TO <lv_tab_part>.
        IF sy-subrc = 0.
          <lv_tab_part> = <lv_comp>.
        ENDIF.
        lv_offset = lv_offset + ls_field-leng.
      ENDLOOP.

      APPEND ls_e071k TO lt_e071k.
    ENDLOOP.

    CALL FUNCTION 'TR_OBJECTS_INSERT'
      EXPORTING
        wi_order                = iv_transport
      TABLES
        wt_ko200                = lt_ko200
        wt_e071k                = lt_e071k
      EXCEPTIONS
        cancel_edit_other_error = 1
        show_only_other_error   = 2
        OTHERS                  = 3.

    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE 'E' NUMBER sy-msgno
        WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.
  ENDMETHOD.

  METHOD file_open_dialog.
    DATA lt_file_table TYPE filetable.
    DATA lv_action TYPE i.
    DATA lv_rc TYPE i.

    cl_gui_frontend_services=>file_open_dialog(
      EXPORTING
        window_title            = 'Select Excel file'
        file_filter             = 'Excel Files (*.xlsx)|*.xlsx|'
      CHANGING
        file_table              = lt_file_table
        rc                      = lv_rc
        user_action             = lv_action
      EXCEPTIONS
        file_open_dialog_failed = 1
        cntl_error              = 2
        error_no_gui            = 3
        not_supported_by_gui    = 4
        OTHERS                  = 5 ).

    IF sy-subrc <> 0 OR lv_action = cl_gui_frontend_services=>action_cancel.
      RETURN.
    ENDIF.

    READ TABLE lt_file_table INDEX 1 INTO DATA(ls_file).
    IF sy-subrc = 0.
      rv_path = ls_file-filename.
    ENDIF.
  ENDMETHOD.

  METHOD file_save_dialog.
    DATA lv_filename TYPE string.
    DATA lv_path TYPE string.
    DATA lv_action TYPE i.

    cl_gui_frontend_services=>file_save_dialog(
      EXPORTING
        window_title         = 'Save Excel file'
        default_extension    = 'xlsx'
        default_file_name    = iv_default_name
        file_filter          = 'Excel Files (*.xlsx)|*.xlsx|'
      CHANGING
        filename             = lv_filename
        path                 = lv_path
        fullpath             = rv_path
        user_action          = lv_action
      EXCEPTIONS
        cntl_error           = 1
        error_no_gui         = 2
        not_supported_by_gui = 3
        OTHERS               = 4 ).

    IF sy-subrc <> 0 OR lv_action = cl_gui_frontend_services=>action_cancel.
      CLEAR rv_path.
    ENDIF.
  ENDMETHOD.
ENDCLASS.
