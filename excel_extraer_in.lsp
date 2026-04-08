;;; --- EXCEL PRO v9.7.1 (IMPERIAL) ---
;;; v9.7.1: Radical Handle Shield, Tag Sanitizer, Feet-Inches.

(vl-load-com)

;; --- FUNCIONES AUXILIARES ---
(defun EX_GEX (n / a r) (setq a "") (while (> n 0) (setq r (rem (1- n) 26) a (strcat (chr (+ 65 r)) a) n (/ (1- n) 26))) a)

(defun EX_GBO (bn / d b r)
  (setq d (vla-get-Blocks (vla-get-ActiveDocument (vlax-get-acad-object))) r '())
  (if (not (vl-catch-all-error-p (setq b (vl-catch-all-apply 'vla-item (list d bn))))) 
    (vlax-for e b (if (= (vla-get-ObjectName e) "AcDbAttributeDefinition") (setq r (cons (vla-get-TagString e) r)))))
  (reverse r)
)

(defun EX_GCA (bn / d b r)
  (setq d (vla-get-Blocks (vla-get-ActiveDocument (vlax-get-acad-object))) r '())
  (if (not (vl-catch-all-error-p (setq b (vl-catch-all-apply 'vla-item (list d bn))))) 
    (vlax-for e b (if (and (= (vla-get-ObjectName e) "AcDbAttributeDefinition") (= (vla-get-Constant e) :vlax-true)) 
      (setq r (cons (cons (vla-get-TagString e) (vla-get-TextString e)) r)))))
  r
)

(defun EX_PIB (tg vl / n q r m pos)
  (setq q (chr 34) r nil n (distof vl))
  ;; --- MOTOR DE FILTRADO v9.7.1 (Anti-Basura Inelectra) ---
  (if (or (member (strcase tg) '("TAG" "TRAMO" "ID" "IDENT" "IDENTIFICADOR" "ID_TRAMO" "ID_TAG" "T_CONDUIT" "HANDLE"))
          (wcmatch (strcase tg) "*TRAMO*") (wcmatch (strcase tg) "*TAG*") (wcmatch (strcase tg) "*ID*"))
    (progn
      ;; Saneamiento: Eliminar comentarios en parentesis y espacios basura
      (if (setq pos (vl-string-search "(" vl)) (setq vl (vl-string-right-trim " " (substr vl 1 pos))))
      (if (setq pos (vl-string-search " " vl)) (setq vl (substr vl 1 pos)))
      
      (setq r (list "TXT" vl))
    )
  )
  
  (if (and (not r) n (/= vl ""))
    (cond 
      ((member (strcase tg) '("LONGITUD" "L" "LEN" "LENGTH" "LONG"))
       (setq m n) (if (< m 39.37) (setq m 39.37) (setq m (* (fix (+ (/ m 39.37) 0.5)) 39.37)))
       (setq r (list "WIN" (strcat (itoa (fix (/ m 12))) "'- " (itoa (rem (fix m) 12)) "\"") m)))
      ((or (wcmatch (strcase tg) "*DIAM*") (wcmatch (strcase tg) "*REDUC*") (member (strcase tg) '("D" "DIA" "DIAMETRO")))
       (setq r (list "TXT" 
         (cond 
           ((equal n 0.75 0.01) "3/4\"")
           ((equal n 1.0 0.01)  "1\"")
           ((equal n 1.5 0.01)  "1 1/2\"")
           ((equal n 2.0 0.01)  "2\"")
           (t (strcat (rtos n 2 2) q))
         ))))
      ((member (strcase tg) '("PZ" "PZS" "CANT" "QTY" "CANTIDAD")) (setq r (list "NUM" (fix (+ n 0.999999)))))
      (t (setq r (list "NUM" n)))
    )
    (if (or (wcmatch (strcase tg) "*DIAM*") (wcmatch vl (strcat "*" q "*"))) (setq r (list "TXT" (strcat "'" vl))) (setq r (list "RAW" vl)))
  ) r
)

;; --- TRADUCTOR DE DIAMETROS ---
(defun EX_DiametroDecimal (dia)
  (cond
    ((= dia "1/2\"") "0.5")
    ((= dia "3/4\"") "0.75")
    ((= dia "1\"") "1")
    ((= dia "1 1/4\"") "1.25")
    ((= dia "1 1/2\"") "1.5")
    ((= dia "2\"") "2")
    ((= dia "2 1/2\"") "2.5")
    ((= dia "3\"") "3")
    ((= dia "4\"") "4")
    (t dia)
  )
)

;; --- LECTOR CSV ROBUSTO (Parseo caracter por caracter) ---
(defun EX_ParseCSVLine (str / len i c inQuotes fld result nextC)
  (setq len (strlen str) i 1 inQuotes nil fld "" result '())
  (while (<= i len)
    (setq c (substr str i 1))
    (if (= c "\"")
      (progn
        (if (< i len) (setq nextC (substr str (1+ i) 1)) (setq nextC ""))
        (if (= nextC "\"")
          (progn
            (setq fld (strcat fld "\"") i (1+ i))
          )
          (progn
            (setq inQuotes (not inQuotes))
          )
        )
      )
      (if (and (= c ",") (not inQuotes))
        (progn
          (setq result (append result (list fld)) fld "")
        )
        (progn
          (setq fld (strcat fld c))
        )
      )
    )
    (setq i (1+ i))
  )
  (setq result (append result (list fld)))
  result
)

;; --- EXTRACCION DE TIPO PARA ACCESORIOS ---
(defun EX_ExtraerTipo (desc / p1 p2)
  (setq p1 (vl-string-search "TIPO \"" desc))
  (if p1
    (progn
      (setq p1 (+ p1 6))
      (setq p2 (vl-string-search "\"" desc p1))
      (if p2
        (substr desc (1+ p1) (- p2 p1))
        "XX"
      )
    )
    "XX"
  )
)

;; --- BUSCAR DENT_CODE EN Book3.csv ---
(defun EX_BuscarDentCode (diametro forma material unidad tipo / csv_path fp line fields col_a col_b fq ft pat_mat pat_forma pat_tipo result found)
  (setq csv_path "C:\\Users\\nleon25050\\Documents\\Antigravity\\Proyectos-de-Inelectra\\Book3.csv")
  (setq fp (open csv_path "r"))
  (if (not fp)
    (progn (princ "\nError: No se puede abrir Book3.csv") (setq result nil))
    (progn
      (setq result nil found nil)
      (setq pat_mat (if (wcmatch (strcase material) "*GALVANIZADO*") "*GALVANIZADO*,*ELECTROGALV*" (strcat "*" (strcase material) "*")))
      (setq pat_forma (strcat "*FORMA " (strcase forma) "*"))
      (setq pat_tipo (if tipo (strcat "*TIPO `\"" (strcase tipo) "*") nil))
      (read-line fp)
      (while (and (not found) (setq line (read-line fp)))
        (if (> (strlen line) 0)
          (progn
            (setq fields (EX_ParseCSVLine line))
            (if (>= (length fields) 20)
              (progn
                (setq col_a (nth 0 fields))
                (setq col_b (nth 1 fields))
                (setq fq (nth 16 fields))
                (setq ft (nth 19 fields))
                (if (and fq ft
                         (= (strcase ft) (strcase unidad))
                         (= (atof fq) (atof diametro))
                         (wcmatch (strcase col_b) "*CONDULETA*")
                         (wcmatch (strcase col_b) pat_mat)
                         (wcmatch (strcase col_b) pat_forma)
                         (or (not pat_tipo) (wcmatch (strcase col_b) pat_tipo))
                    )
                  (progn 
                    (setq result col_a *ultima_desc* col_b found T)
                  )
                )
              )
            )
          )
        )
      )
      (close fp)
    )
  )
  result
)

;; --- MOTOR DE REPORTES (FRAGMENTADO PARA ESTABILIDAD) ---

(defun EX_AnexoTEE (xs cr rs / rt tc)
  (setq rt (vlax-get-property xs 'Range (strcat "B" (itoa cr) ":C" (itoa cr))))
  (vlax-put-property (vlax-get-property xs 'Range (strcat "B" (itoa cr))) 'Value2 "--- ANEXO: ACCESORIOS TEE ---")
  (vlax-put-property (vlax-get-property rt 'Interior) 'Color 13421823)
  (vlax-put-property (vlax-get-property rt 'Font) 'Bold :vlax-true)
  (vlax-put-property (vlax-get-property xs 'Range (strcat "B" (itoa (+ cr 1)))) 'Value2 "Total Extraido")
  (setq tc (vlax-get-property xs 'Range (strcat "C" (itoa (+ cr 1))))) (vlax-put-property tc 'Formula (strcat "=SUBTOTAL(109," rs ")")) (vlax-put-property (vlax-get-property tc 'Font) 'Bold :vlax-true) (vlax-put-property (vlax-get-property tc 'Font) 'Color 16711680)
  (vlax-put-property (vlax-get-property xs 'Range (strcat "B" (itoa (+ cr 2)))) 'Value2 "Calculo 60% TEE")
  (setq tc (vlax-get-property xs 'Range (strcat "C" (itoa (+ cr 2))))) (vlax-put-property tc 'Formula (strcat "=ROUNDUP(C" (itoa (+ cr 1)) "*0.6,0)")) (vlax-put-property (vlax-get-property tc 'Font) 'Bold :vlax-true) (vlax-put-property (vlax-get-property tc 'Font) 'Color 16711680)
  (vlax-put-property (vlax-get-property xs 'Range (strcat "B" (itoa (+ cr 3)))) 'Value2 "Calculo 40% TB")
  (setq tc (vlax-get-property xs 'Range (strcat "C" (itoa (+ cr 3))))) (vlax-put-property tc 'Formula (strcat "=C" (itoa (+ cr 1)) "-C" (itoa (+ cr 2)))) (vlax-put-property (vlax-get-property tc 'Font) 'Bold :vlax-true) (vlax-put-property (vlax-get-property tc 'Font) 'Color 16711680)
  (+ cr 5)
)

(defun EX_AnexoCond (xs cr rs / rt tc)
  (setq rt (vlax-get-property xs 'Range (strcat "B" (itoa cr) ":C" (itoa cr))))
  (vlax-put-property (vlax-get-property xs 'Range (strcat "B" (itoa cr))) 'Value2 "--- ANEXO: CONDULETAS ---")
  (vlax-put-property (vlax-get-property rt 'Interior) 'Color 13421823)
  (vlax-put-property (vlax-get-property rt 'Font) 'Bold :vlax-true)
  (vlax-put-property (vlax-get-property xs 'Range (strcat "B" (itoa (+ cr 1)))) 'Value2 "Total Extraidas")
  (setq tc (vlax-get-property xs 'Range (strcat "C" (itoa (+ cr 1))))) (vlax-put-property tc 'Formula (strcat "=SUBTOTAL(109," rs ")")) (vlax-put-property (vlax-get-property tc 'Font) 'Bold :vlax-true) (vlax-put-property (vlax-get-property tc 'Font) 'Color 16711680)
  (vlax-put-property (vlax-get-property xs 'Range (strcat "B" (itoa (+ cr 2)))) 'Value2 "LR (30%)")
  (setq tc (vlax-get-property xs 'Range (strcat "C" (itoa (+ cr 2))))) (vlax-put-property tc 'Formula (strcat "=ROUND(C" (itoa (+ cr 1)) "*0.3,0)")) (vlax-put-property (vlax-get-property tc 'Font) 'Bold :vlax-true) (vlax-put-property (vlax-get-property tc 'Font) 'Color 16711680)
  (vlax-put-property (vlax-get-property xs 'Range (strcat "B" (itoa (+ cr 3)))) 'Value2 "LB (30%)")
  (setq tc (vlax-get-property xs 'Range (strcat "C" (itoa (+ cr 3))))) (vlax-put-property tc 'Formula (strcat "=ROUND(C" (itoa (+ cr 1)) "*0.3,0)")) (vlax-put-property (vlax-get-property tc 'Font) 'Bold :vlax-true) (vlax-put-property (vlax-get-property tc 'Font) 'Color 16711680)
  (vlax-put-property (vlax-get-property xs 'Range (strcat "B" (itoa (+ cr 4)))) 'Value2 "LL (40%)")
  (setq tc (vlax-get-property xs 'Range (strcat "C" (itoa (+ cr 4))))) (vlax-put-property tc 'Formula (strcat "=C" (itoa (+ cr 1)) "-C" (itoa (+ cr 2)) "-C" (itoa (+ cr 3)))) (vlax-put-property (vlax-get-property tc 'Font) 'Bold :vlax-true) (vlax-put-property (vlax-get-property tc 'Font) 'Color 16711680)
  (+ cr 6)
)

(defun EX_AnexoPipe (xs cr rs mi / rt tc r1)
  (setq rt (vlax-get-property xs 'Range (strcat "B" (itoa cr) ":C" (itoa cr))))
  (vlax-put-property (vlax-get-property xs 'Range (strcat "B" (itoa cr))) 'Value2 "--- ANEXO: CONDUIT ---")
  (vlax-put-property (vlax-get-property rt 'Interior) 'Color 13421823)
  (vlax-put-property (vlax-get-property rt 'Font) 'Bold :vlax-true)
  
  (vlax-put-property (vlax-get-property xs 'Range (strcat "B" (itoa (+ cr 1)))) 'Value2 "Total Pulgadas (IN)")
  (setq r1 (+ cr 1))
  (setq tc (vlax-get-property xs 'Range (strcat "C" (itoa r1)))) (vlax-put-property tc 'Formula (strcat "=SUBTOTAL(109," rs ")")) (vlax-put-property (vlax-get-property tc 'Font) 'Bold :vlax-true) (vlax-put-property (vlax-get-property tc 'Font) 'Color 16711680)
  
  (vlax-put-property (vlax-get-property xs 'Range (strcat "B" (itoa (+ cr 2)))) 'Value2 "Metraje (FT'-IN\")")
  (setq tc (vlax-get-property xs 'Range (strcat "C" (itoa (+ cr 2))))) (vlax-put-property tc 'Formula (strcat "=INT(C" (itoa r1) "/12) & \"'- \" & MOD(C" (itoa r1) ",12) & \"\"\"\"")) (vlax-put-property (vlax-get-property tc 'Font) 'Bold :vlax-true) (vlax-put-property (vlax-get-property tc 'Font) 'Color 16711680)

  (vlax-put-property (vlax-get-property xs 'Range (strcat "B" (itoa (+ cr 3)))) 'Value2 "Metros Totales (MT)")
  (setq tc (vlax-get-property xs 'Range (strcat "C" (itoa (+ cr 3))))) (vlax-put-property tc 'Formula (strcat "=C" (itoa r1) "*0.0254")) (vlax-put-property (vlax-get-property tc 'Font) 'Bold :vlax-true) (vlax-put-property (vlax-get-property tc 'Font) 'Color 16711680)
  
  (vlax-put-property (vlax-get-property xs 'Range (strcat "B" (itoa (+ cr 4)))) 'Value2 "Piezas Estimadas (3mts)")
  (setq tc (vlax-get-property xs 'Range (strcat "C" (itoa (+ cr 4))))) (vlax-put-property tc 'Formula (strcat "=ROUNDUP(C" (itoa (+ cr 3)) "/3,0)")) (vlax-put-property (vlax-get-property tc 'Font) 'Bold :vlax-true) (vlax-put-property (vlax-get-property tc 'Font) 'Color 16711680)
  (+ cr 6)
)

(defun EX_AnexoTipos (xs cr counts forma mat / col_idx d_val total t_lr t_lb t_ll dec_dia rt tc tipo_pair current_assigned target_lr target_lb)
  (setq col_idx 5) ;; Columna E
  (setq tc (vlax-get-property xs 'Range (strcat (EX_GEX col_idx) (itoa cr))))
  (vlax-put-property tc 'Value2 "--- ANEXO: TIPO CONDULETAS ---")
  (vlax-put-property (vlax-get-property tc 'Font) 'Bold :vlax-true)
  (vlax-put-property (vlax-get-property tc 'Interior) 'Color 13421823)
  
  (setq cr (1+ cr))
  (foreach item counts
    (setq d_val (car item) total (cdr item))
    (setq target_lr (fix (+ (* total 0.3) 0.5))
          target_lb (fix (+ (* total 0.3) 0.5))
          t_ll (- total target_lr target_lb))
    (setq dec_dia (EX_DiametroDecimal d_val))
    
    (foreach tipo_pair (list (cons "LR" target_lr) (cons "LB" target_lb) (cons "LL" t_ll))
      (if (> (cdr tipo_pair) 0)
        (progn
          (setq tc (vlax-get-property xs 'Range (strcat (EX_GEX col_idx) (itoa cr))))
          (vlax-put-property tc 'Value2 (strcat (car tipo_pair) " (" d_val ")"))
          (vlax-put-property (vlax-get-property tc 'Font) 'Bold :vlax-true)
          
          (setq tc (vlax-get-property xs 'Range (strcat (EX_GEX col_idx) (itoa (+ cr 1)))))
          (vlax-put-property tc 'Value2 (cdr tipo_pair))
          
          (setq tc (vlax-get-property xs 'Range (strcat (EX_GEX col_idx) (itoa (+ cr 2)))))
          (vlax-put-property tc 'Value2 (EX_BuscarDentCode dec_dia forma mat "IN" (car tipo_pair)))
          
          (setq col_idx (1+ col_idx))
        )
      )
    )
  )
)

;; --- COMANDO PRINCIPAL ---

(defun c:EXCEL_PRO ( / ss i e o n bl current df f id rs sb q c40 c41 xa xb xs r c t_s ty b_n b_o t_g v_l c_i er mi hm ft fc cr h tc ra pz ap range_str spm_col tipo_col dia_val dent_code *forma_idx* *mat_idx* desc_tee cond_counts cond_assigned d_val it_c target_lr target_lb current_lr current_lb current_ll assigned_tipo)
  (princ "\n--- EXCEL PRO v8.8.4 IN ---")
  (if (not *forma_sel*) (setq *forma_sel* "7"))
  (if (not *mat_sel*)   (setq *mat_sel* "Acero Galvanizado"))
  (setq ss (ssget "X" '((0 . "INSERT") (410 . "Model"))))
  (if (not ss) (progn (alert "No hay bloques.") (exit)))
  
  (setq bl '() i 0 q (chr 34) c40 (chr 40) c41 (chr 41))
  (repeat (sslength ss) (setq e (ssname ss i) o (vlax-ename->vla-object e) n (vl-catch-all-apply 'vla-get-EffectiveName (list o))) (if (vl-catch-all-error-p n) (setq n (vla-get-Name o))) (setq n (vl-princ-to-string n)) (if (not (assoc n bl)) (setq bl (cons (cons n 0) bl))) (setq i (1+ i)))
  
  (setq current bl df (strcat (if (getvar "TEMPPREFIX") (getvar "TEMPPREFIX") "C:/Temp/") "ex_in.dcl") f (open df "w"))
  (write-line (strcat "ex : dialog { label=" q "Excel Pro IN v8.5" q "; :column { :edit_box { key=" q "f" q "; label=" q "Filtro:" q "; } :list_box { key=" q "l" q "; multiple_select=true; height=20; width=55; } } :row { :popup_list { key=" q "forma" q "; label=" q "Forma:" q "; width=20; } :popup_list { key=" q "material" q "; label=" q "Material:" q "; width=30; } } :row { :button { key=" q "accept" q "; label=" q "GENERAR" q "; is_default=true; } :button { key=" q "cancel" q "; label=" q "Cerrar" q "; is_cancel=true; } } }") f) (close f)
  
  (defun ULV () (start_list "l") (foreach x current (add_list (strcat (if (= (cdr x) 1) "[X] " "[ ] ") (car x)))) (end_list))
  (defun UF (v / s) (setq s (strcase v) current '()) (foreach x bl (if (wcmatch (strcase (car x)) (strcat "*" s "*")) (setq current (append current (list x))))) (ULV))
  (defun TS (v / idx bn os ns cm it) (setq cm (strcat c40 v c41) idx (read cm)) (foreach i idx (setq bn (car (nth i current)) it (assoc bn bl) os (cdr it) ns (if (= os 1) 0 1) bl (subst (cons bn ns) it bl) current (subst (cons bn ns) (nth i current) current))) (ULV))
  
  (setq id (load_dialog df))
  (if (and id (new_dialog "ex" id))
    (progn (ULV) 
      (start_list "forma") (add_list "7") (add_list "8") (end_list)
      (set_tile "forma" (if (= *forma_sel* "8") "1" "0"))
      (start_list "material") (add_list "Acero Galvanizado") (end_list)
      (set_tile "material" "0")
      (action_tile "f" "(UF $value)") (action_tile "l" "(TS $value)") 
      (action_tile "forma"    "(setq *forma_idx* $value)")
      (action_tile "material" "(setq *mat_idx* $value)")
      (action_tile "accept" "(progn (setq *forma_sel* (if (= *forma_idx* \"1\") \"8\" \"7\")) (setq *mat_sel* \"Acero Galvanizado\") (done_dialog 1))")
      (action_tile "cancel" "(done_dialog 0)") (setq rs (start_dialog)) (unload_dialog id)
      (if (= rs 1)
        (progn (setq sb '()) (foreach x bl (if (= (cdr x) 1) (setq sb (cons (car x) sb))))
          (if sb 
            (progn (setq mi 53 t_s '("BLOQUE") ty '() i 0 hm nil ft 0.0 fc 0.0)
              (repeat (sslength ss) (setq e (ssname ss i) o (vlax-ename->vla-object e) n (vl-princ-to-string (vl-catch-all-apply 'vla-get-EffectiveName (list o)))) (if (and (member n sb) (not (member n ty))) (setq ty (append ty (list n)))) (setq i (1+ i)))
              (foreach b_n ty (setq b_o (EX_GBO b_n)) (foreach t_g b_o (if (not (member t_g t_s)) (setq t_s (append t_s (list t_g))))))
              
              ;; --- Pre-conteo de conduletas para distribucion porcentual ---
              (setq cond_counts '() i 0)
              (repeat (sslength ss)
                (setq e (ssname ss i) o (vlax-ename->vla-object e) n (vl-princ-to-string (vl-catch-all-apply 'vla-get-EffectiveName (list o))))
                (if (and (member n sb) (wcmatch (strcase n) "CONDULETA_*") (not (wcmatch (strcase n) "CONDULETA_TIPO_*")))
                  (progn
                    (setq d_val nil)
                    (if (and (vlax-property-available-p o 'HasAttributes) (= (vla-get-HasAttributes o) :vlax-true))
                      (foreach at (vlax-invoke o 'GetAttributes)
                        (if (member (strcase (vla-get-TagString at)) '("DIAMETRO" "DIAM" "D" "DIA" "INPUT_1")) (setq d_val (vla-get-TextString at)))))
                    (if (not d_val) (foreach co (EX_GCA n) (if (member (strcase (car co)) '("DIAMETRO" "DIAM" "D" "DIA" "INPUT_1")) (setq d_val (cdr co)))))
                    (if d_val
                      (progn
                        (setq it_c (assoc d_val cond_counts))
                        (if it_c (setq cond_counts (subst (cons d_val (1+ (cdr it_c))) it_c cond_counts)) (setq cond_counts (cons (cons d_val 1) cond_counts)))))))
                (setq i (1+ i)))
              (setq cond_assigned '())
              
              ;; --- Agregar columnas "TIPO" y "Codigo SPM" ---
              (setq t_s (append t_s (list "TIPO" "Codigo SPM")))
              (setq spm_col (length t_s) tipo_col (1- spm_col))
              (setq xa (vlax-get-or-create-object "Excel.Application")) (vlax-put-property xa 'Visible :vlax-true) (setq xb (vlax-invoke-method (vlax-get-property xa 'Workbooks) 'Add) xs (vlax-get-property xb 'ActiveSheet))
              (setq c 1) (foreach h t_s (setq tc (vlax-get-property xs 'Range (strcat (EX_GEX c) "1"))) (vlax-put-property tc 'Value2 h) (vlax-put-property (vlax-get-property tc 'Interior) 'Color 6299648) (vlax-put-property (vlax-get-property tc 'Font) 'Color 16777215) (vlax-put-property (vlax-get-property tc 'Font) 'Bold :vlax-true) (setq c (1+ c)))
              (setq r 2 i 0)
              (repeat (sslength ss) 
                (setq e (ssname ss i) o (vlax-ename->vla-object e) n (vl-princ-to-string (vl-catch-all-apply 'vla-get-EffectiveName (list o))))
                (if (member n sb) 
                  (progn 
                    (vl-catch-all-apply 'vlax-put-property (list (vlax-get-property xs 'Range (strcat "A" (itoa r))) 'Value2 n))
                    (setq ra '() pz 1.0) 
                    (if (and (vlax-property-available-p o 'HasAttributes) (= (vla-get-HasAttributes o) :vlax-true)) 
                      (foreach at (vlax-invoke o 'GetAttributes) (setq ra (cons (cons (vla-get-TagString at) (vla-get-TextString at)) ra)))
                    )
                    (foreach co (EX_GCA n) (setq ra (cons co ra)))
                    (foreach ap ra 
                      (setq t_g (car ap) v_l (cdr ap) c_i (vl-position t_g t_s))
                      (if (and c_i (member (strcase t_g) '("PZ" "PZS" "CANT" "QTY" "CANTIDAD"))) 
                        (setq pz (if (and v_l (/= v_l "")) (distof v_l) 1.0))
                      )
                      (if c_i 
                        (progn 
                          (setq er (EX_PIB t_g v_l) tc (vlax-get-property xs 'Range (strcat (EX_GEX (1+ c_i)) (itoa r)))) 
                          (if (not (vl-catch-all-error-p er)) 
                            (cond 
                              ((= (car er) "TXT") (vlax-put-property tc 'Value2 (cadr er))) 
                              ((= (car er) "NUM") (vlax-put-property tc 'Value2 (cadr er))) 
                              ((= (car er) "RAW") (vlax-put-property tc 'Value2 (cadr er))) 
                              ((= (car er) "WIN") 
                                (vlax-put-property tc 'Value2 (cadr er)) 
                                (vlax-put-property (vlax-get-property xs 'Range (strcat (EX_GEX mi) (itoa r))) 'Value2 (caddr er)) 
                                (setq hm T)
                              )
                            )
                          )
                        )
                      )
                    )
                    (if (wcmatch (strcase n) "*CONDULETA_TIPO_TEE*")
                        (progn (vlax-put-property (vlax-get-property xs 'Range (strcat (EX_GEX 54) (itoa r))) 'Value2 pz) (setq ft (+ ft pz)))
                        (if (or (wcmatch (strcase n) "*CONDULETA*") (wcmatch (strcase n) "*LL*") (wcmatch (strcase n) "*LB*") (wcmatch (strcase n) "*LR*"))
                            (progn (vlax-put-property (vlax-get-property xs 'Range (strcat (EX_GEX 55) (itoa r))) 'Value2 pz) (setq fc (+ fc pz)))
                        )
                    )
                    ;; --- Buscar y escribir Codigo SPM y Tipo ---
                    (setq dia_val nil)
                    (foreach ap ra
                      (if (and (not dia_val) (member (strcase (car ap)) '("DIAMETRO" "DIAM" "D" "DIA" "INPUT_1")))
                        (setq dia_val (cdr ap))
                      )
                    )
                    (if (and dia_val (/= dia_val ""))
                      (progn
                        (setq *ultima_desc* nil assigned_tipo nil)
                        (cond 
                          ((wcmatch (strcase n) "*CONDULETA_TIPO_TEE*")
                           (setq assigned_tipo "T"))
                          ((and (wcmatch (strcase n) "CONDULETA_*") (not (wcmatch (strcase n) "CONDULETA_TIPO_*")))
                           (setq it_c (assoc dia_val cond_counts) it_a (assoc dia_val cond_assigned))
                           (if (not it_a) (setq it_a (cons dia_val 0) cond_assigned (cons it_a cond_assigned)))
                           (setq target_lr (fix (+ (* (cdr it_c) 0.3) 0.5))
                                 target_lb (fix (+ (* (cdr it_c) 0.3) 0.5))
                                 current_assigned (cdr it_a))
                           (cond 
                             ((< current_assigned target_lr) (setq assigned_tipo "LR"))
                             ((< current_assigned (+ target_lr target_lb)) (setq assigned_tipo "LB"))
                             (t (setq assigned_tipo "LL"))
                           )
                           (setq cond_assigned (subst (cons dia_val (1+ current_assigned)) it_a cond_assigned))
                           (vlax-put-property (vlax-get-property xs 'Range (strcat (EX_GEX tipo_col) (itoa r))) 'Value2 assigned_tipo))
                        )
                        (setq dent_code (EX_BuscarDentCode (EX_DiametroDecimal dia_val) *forma_sel* *mat_sel* "IN" assigned_tipo))
                        (if dent_code
                          (progn
                            (vlax-put-property (vlax-get-property xs 'Range (strcat (EX_GEX spm_col) (itoa r))) 'Value2 dent_code)
                            (if (and *ultima_desc* (wcmatch (strcase n) "*CONDULETA_TIPO_TEE*")) (setq desc_tee *ultima_desc*))
                          )
                        )
                      )
                    )
                    (setq r (1+ r))
                  )
                )
                (setq i (1+ i))
              )
              (vl-catch-all-apply 'vlax-invoke (list (vlax-get-property (vlax-get-property xs 'Range "A1") 'CurrentRegion) 'AutoFilter))
              (vl-catch-all-apply 'vlax-invoke (list (vlax-get-property xs 'UsedRange) 'Sort (vlax-get-property xs 'Range "A2") 1))
              (setq cr (+ r 2) start_cr cr)
              (if (> ft 0.0) (setq cr (EX_AnexoTEE xs cr (strcat (EX_GEX 54) "2:" (EX_GEX 54) (itoa (1- r))))))
              (setq start_cr cr)
              (if (> fc 0.0) (setq cr (EX_AnexoCond xs cr (strcat (EX_GEX 55) "2:" (EX_GEX 55) (itoa (1- r))))))
              (if (> fc 0.0) (EX_AnexoTipos xs start_cr cond_counts *forma_sel* *mat_sel*))
              (if hm (progn (setq range_str (strcat (EX_GEX mi) "2:" (EX_GEX mi) (itoa (1- r)))) (setq cr (EX_AnexoPipe xs cr range_str mi))))
              
              ;; Forzar actualizacion y expansion total
              (vlax-put-property xa 'ScreenUpdating :vlax-true)
              (vlax-invoke-method (vlax-get-property (vlax-get-property xs 'Cells) 'Columns) 'AutoFit)
              
              ;; Ocultar columnas de calculo con proteccion de error
              (vl-catch-all-apply 'vlax-put-property 
                (list (vlax-get-property xs 'Range (strcat (EX_GEX mi) ":" (EX_GEX 55))) 'Hidden :vlax-true))
              
              (princ "\n>>> BOM GENERADA CON EXPANSIÓN TOTAL (v9.7.0) <<<")
            )
          )
        )
      )
    )
  )
  (if (and df (vl-file-size df)) (vl-file-delete df)) (princ)
)

(princ "\n--- EXCEL PRO v9.7.1 [IMPERIAL PRO] ---") (princ)
