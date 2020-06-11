PACKAGE BODY spr_to_sap

IS
    PROCEDURE start_html
    IS
    BEGIN
    h.p ('<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
        <meta http-equiv="Content-Language" content="ru">');

    h.head;
    HTP.title ('Справочники в Excel для SAP');

    --JS
    HTP.p (jquery12);
    HTP.p (jquery_ui_js);
    HTP.p (jquery_ui_css);
    --css

    HTP.p(css_base);
    h.p('<link rel="stylesheet" href="student.portal_script.h_base_css">');
    h.p('<link rel="stylesheet" href="kadr$.js_gant.gant_css">');

    h.p('<style>
        * {
            margin: 0;
            padding: 0;
            font:13px Helvetica;
        }

        html{
            overflow: hidden;
        }

    </style>
    ');



    END;


    PROCEDURE datepicker
    IS
    BEGIN
        HTP.p ('
    <script type="text/javascript">

            $(document).ready(function () {
            $("#d_b, #d_e").datepicker({
               dateFormat: "dd.mm.yy",
               firstDay:1
               });
        });

    </script>
    ');
    END;


    /*Почему бы не сделать процедуру для логирования?*/
    procedure logirovanie_tables(id_z_1 number,id_sap_old_1 varchar2, id_sap_new_1 varchar2,  date_quit_1 DATE,  name_table_1 varchar2,  name_scheme_1 varchar2, TYPE_CH_1 varchar2)
    is
    syt_change number;
    begin
    insert into personal.spr_log_sap (id_z,id_sap_old, id_sap_new, date_quit, name_table, name_scheme, TYPE_CH)
    VALUES(id_z_1,id_sap_old_1, id_sap_new_1, date_quit_1, name_table_1, name_scheme_1, TYPE_CH_1);

    end;


    /*************************************************************************/
    --Создание XLS
    PROCEDURE setxls(p_date_begin VARCHAR2, p_date_end VARCHAR2, p_select VARCHAR2)
    IS
    name_t varchar(15);
    str    varchar(200);
    strrar  INTEGER :=2;

    coco  INTEGER:=0;
    koko  INTEGER:=0;
    cosum INTEGER:=0;

    type t_rc is ref cursor;
          type t_rec is record(cnt number);
        l_rc t_rc;
          rec t_rec;
        cur_str   varchar(2000);

    BEGIN
    HTP.p ('<script type="text/javascript">
    function jk(){

        var i, j, str,
                ');
    --Создаём справочник по таблицам и чекбоксам
        FOR lp IN(
    SELECT distinct id_table
    FROM sap.ext_tables@kadr$
    WHERE
        pr_upd = 1
    AND id_table !=ALL--не сделаны
                ('PA_054',

                'PA_058',
                'PA_046',
                'BN_020'))
        LOOP

        h.p(lp.id_table||'=document.getElementById("'||lp.id_table||'"),' );
        h.p(lp.id_table||'_c=document.getElementById("'||lp.id_table||'_c"),' );

        END LOOP;

    h.p('
        rowCount,
        oBook = new ActiveXObject("Excel.Application");// Activates Excel
        oBook.Workbooks.Add(); // Opens a new Workbook
        oBook.Caption = "Отчёт";


    var jjo = document.getElementsByTagName("input");
    var date =  new Date();

    var fulldate = date.getFullYear()+""+(date.getMonth()+1)+""+date.getDate();

    //window.alert(fulldate);

    //Строим условия и процедуры
    var array = new Array();
    var nametable = new Array();
    var namesheets = new Array();
    var syt_ch = new Array();
    var date_change = new Array();
    var fio_avtor = new Array();
    ');

           FOR lp IN(
                    SELECT distinct id_table, table_name,s||'.'||t dict
                    FROM sap.ext_tables@kadr$
                    WHERE
                        pr_upd = 1

                   AND id_table !=ALL--не сделаны
                        (
                        'PA_054',
                        'PA_058',
                        'PA_046',
                        'BN_020')

                        )
        LOOP

        h.p('if('||lp.id_table||'){' );
            h.p('if('||lp.id_table||'_c.checked){' );

        h.p('var elements = '||lp.id_table||'.getElementsByTagName("input");');
        str:='begin '||own_name_shapki||'.'||lp.id_table|| '; end;';

                execute immediate str;
        h.p('


                var a=0;

                var stroka='||case when lp.id_table='PA_059' or lp.id_table='PA_042'
                                then '1'
                                when lp.id_table='PA_020' or lp.id_table='PA_035' then '3'

                                    else '2'
                              end||';

                var numbersheets=1;

                var f1=false,f2=false,f3=false;




            for (i = '||case when lp.id_table='PA_059' or lp.id_table='PA_042'
                                then '1'
                            when lp.id_table='PA_020' or lp.id_table='PA_035' then '3'
                            else '2'
                        end||'; i <  '||lp.id_table||'.rows.length; i++) {
                //отмечена ли строка
                if(elements[a].checked==true){
                        //Проверим предыдущая две крайние ячейки предыдущей и текущей строки
                        if(i=='||case
                                    when lp.id_table='PA_059' or lp.id_table='PA_042'
                                        then '1'
                                    when lp.id_table='PA_020' or lp.id_table='PA_035' then '3'
                                        else '2'
                                    end||'){

                                            oSheet.name = "'||lp.id_table||'" +"_"+'||lp.id_table||'.rows[i].cells['||lp.id_table||'.rows[i].cells.length-2].innerText;
                                            namesheets.push("'||lp.id_table||'" +"_"+'||lp.id_table||'.rows[i].cells['||lp.id_table||'.rows[i].cells.length-2].innerText);
                                            date_change.push('||lp.id_table||'.rows[i].cells['||lp.id_table||'.rows[i].cells.length-2].innerText);
                                            syt_ch.push('||lp.id_table||'.rows[i].cells['||lp.id_table||'.rows[i].cells.length-3].innerText);
                                            array.push("'||lp.id_table||'");        //Код справочника
                                            nametable.push("'||lp.table_name||'");  //Наименование справочника
                                            fio_avtor.push('||lp.id_table||'.rows[i].cells['||lp.id_table||'.rows[i].cells.length-1].innerText);


                                    for (j = 0; j < '||lp.id_table||'.rows[i].cells.length-3; j++) {
                                                str = '||lp.id_table||'.rows[i].cells[j].innerText;
                                        if(j>0){
                                            oBook.Worksheets(1).Cells(stroka + 1, j).Value = str; // Writes to the sheet
                                        }
                                    }
                                    stroka++;
                                    
                                    
        }else if( 
        '||lp.id_table||'.rows[i].cells['||lp.id_table||'.rows[i].cells.length-3].innerText=='||lp.id_table||'.rows[i-1].cells['||lp.id_table||'.rows[i].cells.length-3].innerText
        && '||lp.id_table||'.rows[i].cells['||lp.id_table||'.rows[i].cells.length-2].innerText=='||lp.id_table||'.rows[i-1].cells['||lp.id_table||'.rows[i].cells.length-2].innerText
        && '||lp.id_table||'.rows[i].cells['||lp.id_table||'.rows[i].cells.length-1].innerText=='||lp.id_table||'.rows[i-1].cells['||lp.id_table||'.rows[i].cells.length-1].innerText){

        for (j = 0; j < '||lp.id_table||'.rows[i].cells.length-3; j++) {
                                                str = '||lp.id_table||'.rows[i].cells[j].innerText;
                                        if(j>0){
                                            oBook.Worksheets(1).Cells(stroka + 1, j).Value = str; // Writes to the sheet
                                        }
                                    }
        stroka++;
        
        numbersheets=1;
        
        
        
        }else if
        ('||lp.id_table||'.rows[i].cells['||lp.id_table||'.rows[i].cells.length-2].innerText!='||lp.id_table||'.rows[i-1].cells['||lp.id_table||'.rows[i].cells.length-2].innerText){
        
        ');
                            
        execute immediate str;

        h.p('
        stroka=2;                        
            oSheet.name = "'||lp.id_table||'" +"_"+'||lp.id_table||'.rows[i].cells['||lp.id_table||'.rows[i].cells.length-2].innerText;
            namesheets.push("'||lp.id_table||'" +"_"+'||lp.id_table||'.rows[i].cells['||lp.id_table||'.rows[i].cells.length-2].innerText);
            date_change.push('||lp.id_table||'.rows[i].cells['||lp.id_table||'.rows[i].cells.length-2].innerText);
            syt_ch.push('||lp.id_table||'.rows[i].cells['||lp.id_table||'.rows[i].cells.length-3].innerText);
            array.push("'||lp.id_table||'");        //Код справочника
            nametable.push("'||lp.table_name||'");  //Наименование справочника
            fio_avtor.push('||lp.id_table||'.rows[i].cells['||lp.id_table||'.rows[i].cells.length-1].innerText);


            for (j = 0; j < '||lp.id_table||'.rows[i].cells.length-3; j++) {
                        str = '||lp.id_table||'.rows[i].cells[j].innerText;
                if(j>0){
                    oBook.Worksheets(1).Cells(stroka + 1, j).Value = str; // Writes to the sheet
                }
            }
            stroka++;
                                    }else{');
                            
                            execute immediate str;

                            h.p('
                            
                            if(numbersheets==1){
                                oSheet.name = "'||lp.id_table||'" +"_"+ '||lp.id_table||'.rows[i].cells['||lp.id_table||'.rows[i].cells.length-2].innerText;
                                namesheets.push("'||lp.id_table||'" +"_"+ '||lp.id_table||'.rows[i].cells['||lp.id_table||'.rows[i].cells.length-2].innerText);
                            
                            }else{
                            
                                    oSheet.name = "'||lp.id_table||'" +"_"+ '||lp.id_table||'.rows[i].cells['||lp.id_table||'.rows[i].cells.length-2].innerText+"_"+numbersheets;
                                    namesheets.push("'||lp.id_table||'" +"_"+ '||lp.id_table||'.rows[i].cells['||lp.id_table||'.rows[i].cells.length-2].innerText+"_"+numbersheets);
                            }  
                                    
                            
                                    syt_ch.push('||lp.id_table||'.rows[i].cells['||lp.id_table||'.rows[i].cells.length-3].innerText);
                                    date_change.push('||lp.id_table||'.rows[i].cells['||lp.id_table||'.rows[i].cells.length-2].innerText);
                                    array.push("'||lp.id_table||'");        //Код справочника
                                    nametable.push("'||lp.table_name||'");  //Наименование справочника
                                    fio_avtor.push('||lp.id_table||'.rows[i].cells['||lp.id_table||'.rows[i].cells.length-1].innerText);
                            
                            
                                    stroka='||case when lp.id_table='PA_059'
                                                then '1'
                                                when lp.id_table='PA_020' or lp.id_table='PA_035' then '3'
                                              else '2'
                                              end||';

                                    for (j = 0; j < '||lp.id_table||'.rows[i].cells.length-3; j++) {
                                                str = '||lp.id_table||'.rows[i].cells[j].innerText;
                                        if(j>0){
                                            oBook.Worksheets(1).Cells(stroka + 1, j).Value = str; // Writes to the sheet
                                        }
                                    }
                                stroka++;numbersheets++;

                            }
                            

                }a++;
            }



        ');

            h.p('}');
        h.p('}');

        end loop;

    h.p('
    //Хотя бы один есть выбранный чекбокс
    var abc = 0;
    if(jjo.length>0){
    for(var k=0; k<jjo.length; k++){
        if((jjo[k].type == "checkbox")
            && (jjo[k].checked == true)){


                abc++;
            }
        }
    }
    //Тогда показываем сформированный EXCEL
    if(abc>0){
    //Если у нас есть таблицы формируем лист "Список изменений"
        oBook.Sheets.Add();
        oSheet = oBook.ActiveSheet;
            oSheet.name = "Список изменений";



        oBook.Worksheets(1).Range("A1").Value = "Код справочника";
        oBook.Worksheets(1).Range("B1").Value = "Наименование справочника";
        oBook.Worksheets(1).Range("C1").Value = "Источник изменения";
        oBook.Worksheets(1).Range("D1").Value = "Суть изменения";
        oBook.Worksheets(1).Range("E1").Value = "Ссылка на изменения";
        oBook.Worksheets(1).Range("F1").Value = "Дата внесения изменения";
        oBook.Worksheets(1).Range("G1").Value = "Автор изменения";
        oBook.Worksheets(1).Range("H1").Value = "Комментарий автора";
        oBook.Worksheets(1).Range("A1:H1").Font.Bold = true;

    /******/

        oBook.Worksheets(1).Range("A1:H1").WrapText=true;
        oSheet.Range("A1:H1").HorizontalAlignment = -4108;

    /*****/
    var bbn=2;
        for(var k=0; k<namesheets.length; k++){

                    oBook.Worksheets(1).Range("A"+bbn).Value = array[k] ;            //Код справочника
                    oBook.Worksheets(1).Range("B"+bbn).Value = nametable[k];        //Наименование справочника
                    oBook.Worksheets(1).Range("C"+bbn).Value = "Внешние системы";   //Источник изменения
                    oBook.Worksheets(1).Range("D"+bbn).Value = syt_ch[k];           //Суть изменения

                    oBook.Worksheets(1).Range("E"+bbn).Value =  namesheets[k];      //Наименование листа страницы

                    oBook.Worksheets(1).Range("F"+bbn).Value = date_change[k];      //Дата внесения изменения
                    oBook.Worksheets(1).Range("G"+bbn).Value = fio_avtor[k];        //Автор изменения


                bbn++;



        }
        oSheet.Columns("A:W").AutoFit;
        oBook.Application.Visible = true;
        //Закрываем отсылки, чтобы не плодились в ДЗ
            oBook = null;
        }
        return;
    }
    </script>

    ');
    exception
    WHEN OTHERS
        THEN
    h.br;
            h.p (SQLERRM);
    END;


    PROCEDURE javascripts
    is
    begin
    h.p('
    <script type="text/javascript">
    function vk(obj)
    {
        var f = document.getElementById(obj);
        if(f.style.display=="block")
             f.style.display="none";
        else f.style.display="block";
    }


    function jj(tableid,obj)
    {
    var     ch = document.getElementById(obj);
    var     table = document.getElementById(tableid);
    var     elements = table.getElementsByTagName("input");

    var     stringt;

        if(ch.checked){


            for(var i=0; i<elements.length; i++)
            {
                elements[i].checked=true;

            }


        }else{

            for(var i=0; i<elements.length; i++)
            {
                elements[i].checked=false;
            }
        }

    }

    function ll(tableid, obj)
    {
    var     ch = document.getElementById(obj);
    var     table = document.getElementById(tableid);
    var     elements = table.getElementsByTagName("input");

    var a=0;

        for(var i=0; i<elements.length; i++)
        {
            if(elements[i].checked==true)
            {a++;}
        }

        if(a>0){
            ch.checked=true;

        }else{
            ch.checked=false;
        }




    }

    function kk(obj,label)
    {
    var     checkbox = document.getElementById(obj);
    var     labelcheck = document.getElementById(label);

    var     chs = document.getElementsByTagName("input");


        if(checkbox.checked==true)
        {
        labelcheck.innerHTML="Снять всё";
            for(var i=0; i<chs.length; i++)
            {
                var chch = chs[i];
                if(chch.getAttribute("type")=="checkbox")
                {
                    if(chs[i].checked==false){
                    chs[i].checked = true;
                    }
                }
            }


            //document.write(chch.getAttribute("type"));
        }else{
        labelcheck.innerHTML="Отметить всё";
            for(var i=0; i<chs.length; i++)
            {
                var chch = chs[i];
                if(chch.getAttribute("type")=="checkbox")
                {
                    if(chs[i].checked==true){
                    chs[i].checked = false;
                    }
                }
            }

        }

        //window.alert(labelcheck.value);
    }


    function hidde()
    {
    var     chs = document.getElementsByTagName("table");

        for(var i=0; i<chs.length; i++)
        {
            if(chs[i].style.display=="block"){

                chs[i].style.display="none";
            }


        }
    window.alert(kk);

    }
    </script>

    ');
    end;


    --Для checkbox и link
    procedure checkboxlink(names varchar2)
    is
    valuecl varchar(200);
    begin

    SELECT distinct table_name
    INTO valuecl
    FROM sap.ext_tables@kadr$
    WHERE
        pr_upd = 1
    AND id_table = names
    AND id_table !=ALL--не сделаны
                (
                'PA_046'
                );
    h.br;
    HTP.formCheckbox (

                    cname          => names||'_c',
                    cvalue         => valuecl,
                    cattributes   =>    ' id="'||names||'_c'||'" class="check" data-name="'||names||'" style="text-decoration: none;" onclick=jj("'||names||'","'||names||'_c'||'")
                     ');

    HTP.anchor (
                    curl          => '#',
                    ctext         => valuecl,
                    cattributes   =>    ' style="text-decoration: none" onclick=vk("'
                                     || names
                                     || '") ');
    end;
    --Вывод HTML разметки******************************************************
    PROCEDURE view_tables (p_date_begin    VARCHAR2 DEFAULT NULL,
                           p_date_end      VARCHAR2 DEFAULT NULL,
                           p_select        VARCHAR2 DEFAULT NULL)
    IS


        ----------------------------------------
        v_date_begin      DATE;
        v_date_end        DATE;
        pp_select         VARCHAR (500) := '';
        t_json kadr$soc.json;


        type t_rc is ref cursor;
          type t_rec is record(cnt number);
        l_rc t_rc;
          rec t_rec;
        cur_str   varchar(2000);
        p_to_select varchar (1);

        ss varchar(500);
        checkboxss varchar(500);
        ----------------------------------------
    BEGIN
    --Проверка-----------------------------------------------------
        IF p_date_begin IS NULL
        THEN
            v_date_begin :=
                TO_DATE (
                    EXTRACT (DAY FROM SYSDATE)
                    || '.'
                    || EXTRACT (MONTH FROM SYSDATE)
                    || '.'
                    || EXTRACT (YEAR FROM SYSDATE),
                    'dd.mm.yyyy');
        ELSE
            v_date_begin := TO_DATE (p_date_begin, 'dd.mm.yyyy');
        END IF;

        IF p_date_end IS NULL
        THEN
            v_date_end :=
                TO_DATE (
                    EXTRACT (DAY FROM SYSDATE)
                    || '.'
                    || EXTRACT (MONTH FROM SYSDATE)
                    || '.'
                    || EXTRACT (YEAR FROM SYSDATE),
                    'dd.mm.yyyy');
        ELSE
            v_date_end := TO_DATE (p_date_end, 'dd.mm.yyyy');
        END IF;


        IF p_select is null
        then
            p_to_select:='1';
        else
            p_to_select := p_select;
        end if;

        for r in 1..3 loop
            pp_select := pp_select || '<option value="'|| r ||'"'|| h.iif(nvl(p_to_select, 0) = r, ' selected', '') ||'>' || case when r = 1 then 'Дате открытия' when r = 2 then 'Дате редактирования' else 'Дате открытия и редактирования' end ||'</option>';
        end loop;

    --Шапочка------------------------------------------------------
    h.html;
        start_html;
        datepicker;
        setxls(p_date_begin, p_date_end, p_select);
        javascripts;
    h.headc;
    h.body;
    --Параметры---------------------------------------------------
    h.p('<div id="filterBG">
               <div id="filterHeader"></div>
               <div id="filterContent">
                   <div class="option">');
    HTP.p (
        '<form method="get" action="'|| get_zapros || '">
                <select name="p_select">'|| pp_select||'
                </select>
                с <input type="text" id="d_b" name="p_date_begin" size="10" value="'
                    || TO_CHAR (v_date_begin, 'dd.mm.yyyy')
                    || '">
                по <input type="text" id="d_e" name="p_date_end" size="10" value="'
                    || TO_CHAR (v_date_end, 'dd.mm.yyyy')
                    || '">
            </div>
                <input type="submit" value="Смотреть" style="width:120px;" >
            </form>
        </div>
    </div>
    ');
    -----------------------------------------------------------------------
    --Сформировать EXCEL---------------------------------------------------
        HTP.p (
            '<input type="submit" id="bb" onclick=jk(); value="Сформировать EXCEL"/>');
        HTP.p (
            '<input type="button" onclick=hidde(); value="Скрыть таблицы">');
    -----------------------------------------------------------------------
    --Справочники----------------------------------------------------------
        HTP.hr;
        HTP.p (
               '<input type="checkbox" id="foralltable" onclick=kk("foralltable","labelforlable"); />
               <label id="labelforlable">Отметить все</label>
               ');


        htp.hr(cattributes=> 'noshade  align="left" size="2" style=" width:350px"');
        HTP.p ('<b>Справочники');

        IF v_date_end IS NULL
        THEN
            HTP.p ('&nbsp;на&nbsp;' || TO_CHAR(v_date_begin,'dd.mm.yyyy'));
        ELSE
            HTP.p (
                   '&nbsp;за период&nbsp;'
                || TO_CHAR(v_date_begin,'dd.mm.yyyy')
                || '-'
                || TO_CHAR(v_date_end, 'dd.mm.yyyy'));
        END IF;

        HTP.p ('</b>');

    /*****************Мощная штука******************************************/
        for r in (
            select distinct id_table infotype, s||'.'||t dict from sap.ext_tables@kadr$ where pr_upd = 1 --AND id_table='PA_032''PA_035'
             AND id_table = ANY( 'PA_032', 'PA_059', 'BN_011', 'PA_013','PA_057','PA_054','PA_053','PA_022', 'PA_024','PA_026', 'PA_048', 'PA_042','PA_028', 'PA_002','OM_002','PA_011', 'OM_005'
             ,'OM_008','PA_004','PA_005','PA_006','PA_012','PA_010','PA_020','PA_035','PA_058')
            /*AND id_table !=ALL--не сделаны
                (
                'PA_054',
                'PA_035',
                'PA_046',
                'PA_058')*/
        ) loop


        cur_str := 'select  count(*) from '|| r.dict ||' where '||
        case
        ------------------------------------------------------------------------------------------------------------------------------------
        when nvl(p_to_select, 0)=1 then
        --у некоторых есть dateenter or date_e придётся подстроиться под них увы :(
            case
            when nvl(r.dict,0)='KADR$.SP_LG_OTP'
                then 'date_e between to_date('''||p_date_begin||''',''dd.mm.yyyy'') AND to_date('''||p_date_end||''',''dd.mm.yyyy'')'
            when
                 nvl(r.dict,0)='Personal.SP_TYPE_PRODUCTION'
                 then 'dateenter between to_date('''||p_date_begin||''',''dd.mm.yyyy'') AND to_date('''||p_date_end||''',''dd.mm.yyyy'')'
            --У таких справочников как:
            -->>>>Вид документа к награде
            -->>>>>Вид документа об образовании
            -->>>>>Вид документа работника
            --Есть условие по MGR
            when nvl(r.infotype,0)='PA_006'
                then 'date_enter between to_date('''||p_date_begin||''',''dd.mm.yyyy'') AND to_date('''||p_date_end||''',''dd.mm.yyyy'') AND MGR=ANY(1,3,5,6,7,11,12) AND id_sap!=99999999'
            when nvl(r.infotype,0)='PA_005'
                then 'date_enter between to_date('''||p_date_begin||''',''dd.mm.yyyy'') AND to_date('''||p_date_end||''',''dd.mm.yyyy'') AND MGR=2 AND id_sap!=99999999'
            when nvl(r.infotype,0)='PA_004'
                then 'date_enter between to_date('''||p_date_begin||''',''dd.mm.yyyy'') AND to_date('''||p_date_end||''',''dd.mm.yyyy'') AND MGR=ANY(9,10) AND id_sap!=99999999'

            else
            'date_enter between to_date('''||p_date_begin||''',''dd.mm.yyyy'') AND to_date('''||p_date_end||''',''dd.mm.yyyy'')'
            end
        ------------------------------------------------------------------------------------------------------------------------------------
        when nvl(p_to_select, 0)=2 then
        'date_quit between to_date('''||p_date_begin||''',''dd.mm.yyyy'') AND to_date('''||p_date_end||''',''dd.mm.yyyy'')'
        ------------------------------------------------------------------------------------------------------------------------------------
        else
            case
            when nvl(r.dict,0)='KADR$.SP_LG_OTP'
                then  'date_e between to_date('''||p_date_begin||''',''dd.mm.yyyy'') AND to_date('''||p_date_end||''',''dd.mm.yyyy'') OR date_quit between to_date('''||p_date_begin||''',''dd.mm.yyyy'') AND to_date('''||p_date_end||''',''dd.mm.yyyy'')'
            when nvl(r.dict,0)='Personal.SP_TYPE_PRODUCTION'
                then 'dateenter between to_date('''||p_date_begin||''',''dd.mm.yyyy'') AND to_date('''||p_date_end||''',''dd.mm.yyyy'') OR date_quit between to_date('''||p_date_begin||''',''dd.mm.yyyy'') AND to_date('''||p_date_end||''',''dd.mm.yyyy'')'
            when nvl(r.infotype,0)='PA_006'
                then 'date_enter between to_date('''||p_date_begin||''',''dd.mm.yyyy'') AND to_date('''||p_date_end||''',''dd.mm.yyyy'') AND MGR=ANY(1,3,5,6,7,11,12) AND id_sap!=99999999 OR date_quit between to_date('''||p_date_begin||''',''dd.mm.yyyy'') AND to_date('''||p_date_end||''',''dd.mm.yyyy'')'
            when nvl(r.infotype,0)='PA_005'
                then 'date_enter between to_date('''||p_date_begin||''',''dd.mm.yyyy'') AND to_date('''||p_date_end||''',''dd.mm.yyyy'') AND MGR=2 AND id_sap!=99999999 OR date_quit between to_date('''||p_date_begin||''',''dd.mm.yyyy'') AND to_date('''||p_date_end||''',''dd.mm.yyyy'')'
            when nvl(r.infotype,0)='PA_004'
                then 'date_enter between to_date('''||p_date_begin||''',''dd.mm.yyyy'') AND to_date('''||p_date_end||''',''dd.mm.yyyy'') AND MGR=ANY(9,10) AND id_sap!=99999999 OR date_quit between to_date('''||p_date_begin||''',''dd.mm.yyyy'') AND to_date('''||p_date_end||''',''dd.mm.yyyy'')'


            else
            'date_enter between to_date('''||p_date_begin||''',''dd.mm.yyyy'') AND to_date('''||p_date_end||''',''dd.mm.yyyy'') OR date_quit between to_date('''||p_date_begin||''',''dd.mm.yyyy'') AND to_date('''||p_date_end||''',''dd.mm.yyyy'')'
            end
        ------------------------------------------------------------------------------------------------------------------------------------
        end;

        --h.p(cur_str);

        open l_rc for cur_str;
            loop
                fetch l_rc into rec;
                exit when l_rc%notfound;
            end loop;
        close l_rc;
            if (rec.cnt>0) then
                ss := 'begin '||own_name_www||'.'||r.infotype||'_html('''||to_char(v_date_begin,'dd.mm.yyyy')||''','''||to_char(v_date_end,'dd.mm.yyyy')||''','''||p_to_select||''','''||r.dict||'''); end;';

                --h.p(cur_str);

                checkboxss := 'begin '||own_name||'.checkboxlink('''||r.infotype||'''); end;';
                execute immediate checkboxss;h.p(rec.cnt); if(r.infotype='PA_058') then h.p('waiting'); end if;

                execute immediate ss;
                --h.p(cur_str);

            end if;

        end loop;

    htp.hr();
    h.bodyc;
    h.htmlc;
    exception
    WHEN OTHERS
        THEN
    htp.br;
            HTP.p (SQLERRM);
    END;
    /*************************************************************************/

END;
Одно из представлений процедуры разметки на веб странице:
--Льготы на авиаперелет
    procedure pa_020_html(p_date_begin    VARCHAR2 DEFAULT NULL,
                          p_date_end      VARCHAR2 DEFAULT NULL,
                          p_select        VARCHAR2 DEFAULT NULL,
                          real_name_table VARCHAR2 DEFAULT NULL)
        is
            ----------------------------------------
            up_part1          INTEGER;
            up_part2          INTEGER;
            row_span          INTEGER;
            ---------------------------------------
            st_part1          VARCHAR (250);
            st_part1_1        VARCHAR (50);
            st_part1_2        VARCHAR (50);
            st_part1_3        VARCHAR (50);
            st_part1_4        VARCHAR (50);
            st_part1_5        VARCHAR (50);
            st_part1_6        VARCHAR (50);
            st_part1_7        VARCHAR (50);
            ----------------------------------------
            st_part2          VARCHAR (250);
            st_part2_1        VARCHAR (50);
            st_part2_2        VARCHAR (50);
            ----------------------------------------
            st_part2_1_1        VARCHAR (50);
            st_part2_2_2        VARCHAR (50);
            ----------------------------------------
            v_date_begin      DATE;
            v_date_end        DATE;
            pp_select         VARCHAR (500);
            checkboxes        INTEGER;
            ----------------------------------------
            name_table_show   VARCHAR (30);
        begin

        v_date_begin := TO_DATE (p_date_begin, 'dd.mm.yyyy');
        v_date_end := TO_DATE (p_date_end, 'dd.mm.yyyy');

        name_table_show := 'PA_020';
--Переменные для объдинения ячек-------------------------------------------------------------------------
                st_part1 := 'АС "Персонал"';
                st_part2 := 'SAP';
        --Часть1-------------------------------------------------------------------------------------------------
                up_part1 := 7;


                st_part1_1 := 'ID';
                st_part1_2 := 'ID_P';
                st_part1_3 := 'NAME';
                st_part1_4 :='FULL_NAME';
                st_part1_5 :='COMM';
                st_part1_6 :='DATE_B';
                st_part1_7 :='DATE_E';

                row_span := 2;
        --Часть2-------------------------------------------------------------------------------------------------
                up_part2 := 2;
                st_part2_1 := 'Код';
                st_part2_2 := 'Текст';

                st_part2_1_1 := 'ZZAIRL';
                st_part2_2_2 := 'ZZAIRL_TXT';
---------------------------------------------------------------------------------------------------------
--Выборка------------------------------------------------------------------------------------------------
                HTP.tableopen (
                    cborder       => 'BORDER=1',
                    cattributes   =>    'bordercolor="#EAD4C1" width="70%" id='
                                     || name_table_show
                                     || ' style="display:none;"',
                    calign        => 'CENTER');
                HTP.tablerowopen;
                HTP.tabledata (cvalue        => st_part1,
                               cattributes   => 'colspan=' || up_part1,
                               calign        => 'CENTER');

                HTP.tabledata (cvalue        => st_part2,
                               cattributes   => 'colspan=' || up_part2,
                               calign        => 'CENTER');



------------------------------------------------------------------------
                HTP.tablerowclose;
                --
                HTP.tablerowopen;

                HTP.tabledata (cvalue => st_part1_1, calign => 'CENTER',
                                cattributes   => 'rowspan=' || row_span);
                HTP.tabledata (cvalue => st_part1_2, calign => 'CENTER',
                                cattributes   => 'rowspan=' || row_span);
                HTP.tabledata (cvalue => st_part1_3, calign => 'CENTER',
                                cattributes   => 'rowspan=' || row_span);
                HTP.tabledata (cvalue => st_part1_4, calign => 'CENTER',
                                cattributes   => 'rowspan=' || row_span);
                HTP.tabledata (cvalue => st_part1_5, calign => 'CENTER',
                                cattributes   => 'rowspan=' || row_span);
                HTP.tabledata (cvalue => st_part1_6, calign => 'CENTER',
                                cattributes   => 'rowspan=' || row_span);
                HTP.tabledata (cvalue => st_part1_7, calign => 'CENTER',
                                cattributes   => 'rowspan=' || row_span);



                HTP.tabledata (cvalue => st_part2_1, calign => 'CENTER');
                HTP.tabledata (cvalue => st_part2_2, calign => 'CENTER');
        htp.tableRowClose;

        htp.tableRowOpen;
                HTP.tabledata (cvalue => st_part2_1_1, calign => 'CENTER');
                HTP.tabledata (cvalue => st_part2_2_2, calign => 'CENTER');

        HTP.tablerowclose;



                    FOR lp
                        IN ((SELECT a.ID
                                ,a.ID_P
                                ,a.NAME
                                ,a.FULL_NAME
                                ,a.COMM
                                ,a.DATE_B
                                ,a.DATE_E,
                               (case
                                when p_select = 1
                                 then a.date_e
                                when p_select = 2
                                 then a.date_quit
                                when p_select = 3 and a.date_e between v_date_begin and v_date_end
                                 then a.date_e
                                when p_select = 3 and a.date_quit between v_date_begin and v_date_end
                                 then a.date_quit
                                else null
                              end) date_g,
                              (case
                                   when n.id_sap_old is null and n.id_sap_new is null
                                     then 'добавление позиции'
                                   when n.id_sap_old is null and n.id_sap_new is not null
                                     then 'добавление мэппинга'
                                   when n.id_sap_old is not null and n.id_sap_new is null
                                     then 'удаление мэппинга'
                                   when n.id_sap_old is not null and n.id_sap_new is not null and  n.id_sap_old <> n.id_sap_new
                                     then 'корректировка мэппинга'

                                   when nvl(n.id_sap_old,'') = nvl(n.id_sap_new,'')
                                     then 'корректировка позиции'

                              end) type_chn,

                            (case
                                when terminal_quit is not null
                                    then  kadr$do.orders_1c_fio.fio_ini(terminal_quit, sysdate)
                                when terminal_enter is not null
                                    then  kadr$do.orders_1c_fio.fio_ini(terminal_enter, sysdate)
                                else '' end) terminal

                          FROM KADR$.SP_LG_OTP a, personal.spr_log_sap n
                             where ((a.date_e between v_date_begin AND v_date_end) and p_select = 1 AND (a.id=n.id_z
                                            and (a.date_e=n.date_quit or  a.date_quit=n.date_quit)
                                            and (n.date_quit between  v_date_begin and v_date_end)
                                            and lower(real_name_table)=lower(n.name_scheme||'.'||n.name_table)
                                            ))
                                or ((a.date_quit  between v_date_begin AND v_date_end) and p_select = 2 AND (a.id=n.id_z
                                            and (a.date_e=n.date_quit or  a.date_quit=n.date_quit)
                                            and (n.date_quit between  v_date_begin and v_date_end)
                                            and lower(real_name_table)=lower(n.name_scheme||'.'||n.name_table)))

                                or ( a.id=n.id_z
                                            AND (a.date_e=n.date_quit OR  a.date_quit=n.date_quit)
                                            AND (n.date_quit between  v_date_begin AND v_date_end)
                                            AND (a.date_e between v_date_begin AND v_date_end OR a.date_quit  between v_date_begin AND v_date_end )
                                            AND LOWER(real_name_table)=LOWER(n.name_scheme||'.'||n.name_table)
                                            AND p_select = 3)
                            UNION
                          SELECT ID
                                ,ID_P
                                ,NAME
                                ,FULL_NAME
                                ,COMM
                                ,DATE_B
                                ,DATE_E,
                                   (case
                                when p_select = 1
                                 then date_e
                                when p_select = 2
                                 then date_quit
                                when p_select = 3 and date_e between v_date_begin and v_date_end
                                 then date_e
                                when p_select = 3 and date_quit between v_date_begin and v_date_end
                                 then date_quit
                                else null
                              end) date_g,
                              (case
                                when date_e between v_date_begin and v_date_end
                                 then 'добавление позиции'
                                when date_quit between v_date_begin and v_date_end
                                 then 'корректировка позиции'
                              end) type_chn,

                              (case
                                when terminal_quit is not null
                                    then  kadr$do.orders_1c_fio.fio_ini(terminal_quit, sysdate)
                                when terminal_enter is not null
                                    then  kadr$do.orders_1c_fio.fio_ini(terminal_enter, sysdate)
                                else '' end) terminal



                          FROM KADR$.SP_LG_OTP
                             where ((date_e between v_date_begin AND v_date_end) and p_select = 1 AND (id not in (SELECT jj.id_z FROM personal.spr_log_sap jj WHERE jj.date_quit between  v_date_begin AND v_date_end)))
                                or ((date_quit  between v_date_begin AND v_date_end) and p_select = 2 AND (id not in (SELECT jj.id_z FROM personal.spr_log_sap jj WHERE jj.date_quit between  v_date_begin AND v_date_end)))
                                or (  (date_e between v_date_begin AND v_date_end OR date_quit  between v_date_begin AND v_date_end )
                                             and p_select = 3 AND (id not in (SELECT jj.id_z FROM personal.spr_log_sap jj WHERE jj.date_quit between  v_date_begin AND v_date_end)))
									)
                                order by type_chn, date_g, terminal ASC)
                    LOOP
                    HTP.tablerowopen;

                     HTP.tabledata(cvalue=>
                        HTF.formCheckbox (

                        cname          => '',
                        cvalue         => '',
                        cattributes   =>    ' style="text-decoration: none" id="'||name_table_show||'_c_'||checkboxes||'"
                        onclick=ll("'||name_table_show||'","'||name_table_show||'_c'||'") ')
                        );
                        HTP.tabledata (cvalue => lp.id, calign => 'LEFT');
                        HTP.tabledata (cvalue => lp.id_p, calign => 'LEFT');
                        HTP.tabledata (cvalue => lp.name, calign => 'LEFT');
                        HTP.tabledata (cvalue => lp.full_name, calign => 'LEFT');
                        HTP.tabledata (cvalue => lp.comm, calign => 'LEFT');
                        HTP.tabledata (cvalue => lp.date_b, calign => 'LEFT');
                        HTP.tabledata (cvalue => lp.date_e, calign => 'LEFT');

                        HTP.tabledata (cvalue => '&nbsp', calign => 'LEFT');
                        HTP.tabledata (cvalue => '&nbsp', calign => 'LEFT');
                        
                        HTP.tabledata (cvalue => lp.type_chn, calign => 'LEFT');
			            HTP.tabledata (cvalue => TO_CHAR(lp.date_g, 'dd.mm.yyyy'), calign => 'LEFT');
			            HTP.tabledata (cvalue => lp.terminal, calign => 'LEFT');

                    HTP.tablerowclose;
                    END LOOP;

        HTP.tableclose;
        END;
Одно из представлений описания шапки для Excel:
--Льготы на авиаперелет
    procedure pa_020
    is
    begin
        HTP.p (
            '

    var elements = PA_020.getElementsByTagName("input");

    oBook.Sheets.Add();
         oSheet = oBook.ActiveSheet;
         oSheet.name = "Льготы на авиаперелет";

    oBook.Worksheets(1).Range("A1:E1").Borders.LineStyle = true;
    oBook.Worksheets(1).Range("A2:E2").Borders.LineStyle = true;

    oBook.Worksheets(1).Range("A1").Value = "АС Персонал";

    oBook.Worksheets(1).Range("A2").Value = "ID";
    oBook.Worksheets(1).Range("B2").Value = "ID_P";
    oBook.Worksheets(1).Range("C2").Value = "NAME";
    oBook.Worksheets(1).Range("D2").Value = "FULL_NAME";
    oBook.Worksheets(1).Range("E2").Value = "COMM";
    oBook.Worksheets(1).Range("F2").Value = "DATE_B";
    oBook.Worksheets(1).Range("G2").Value = "DATE_E";


    oBook.Worksheets(1).Range("A2:A3").MergeCells = true;
    oBook.Worksheets(1).Range("B2:B3").MergeCells = true;
    oBook.Worksheets(1).Range("C2:C3").MergeCells = true;
    oBook.Worksheets(1).Range("D2:D3").MergeCells = true;
    oBook.Worksheets(1).Range("E2:E3").MergeCells = true;
    oBook.Worksheets(1).Range("F2:F3").MergeCells = true;
    oBook.Worksheets(1).Range("G2:G3").MergeCells = true;

    oBook.Worksheets(1).Range("H1").Value = "SAP";

    oBook.Worksheets(1).Range("H2").Value = "Код";
    oBook.Worksheets(1).Range("H3").Value = "ZZAIRL";

    oBook.Worksheets(1).Range("I2").Value = "Текст";
    oBook.Worksheets(1).Range("I3").Value = "ZZAIRL_TXT";

    oBook.Worksheets(1).Range("H1:I1").MergeCells = true;


    oSheet.Columns("A:W").WrapText=false;
    oSheet.Range("A1:W1").Font.Bold = true;
    oSheet.Cells.VerticalAlignment = -4108;
    oSheet.Range("A1:W1").HorizontalAlignment = -4108;


    oSheet.Columns("A:W").AutoFit;
    oBook.Worksheets(1).Range("A1:E1").Interior.Color = 12632256;
    oBook.Worksheets(1).Range("A2:E2").Interior.Color = 12632256;
    oBook.Worksheets(1).Range("A1:E1").Font.Bold = true;
    oBook.Worksheets(1).Range("A2:E2").Font.Bold = true;
    ');
    end;
