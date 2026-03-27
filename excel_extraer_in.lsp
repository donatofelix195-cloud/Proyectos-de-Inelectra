;;; --- EXCEL PRO v9.6.0 (IMPERIAL) ---
;;; v9.6.0: ScreenUpdating Fix, AutoFit Cells, Reductores PRO.

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

(defun EX_PIB (tg vl / n q r m)
  (setq q (chr 34) r nil n (distof vl))
  (if (and n (/= vl ""))
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

;; --- COMANDO PRINCIPAL ---

(defun c:EXCEL_PRO ( / ss i e o n bl current df f id rs sb q c40 c41 xa xb xs r c t_s ty b_n b_o t_g v_l c_i er mi hm ft fc cr h tc ra pz ap range_str)
  (princ "\n--- EXCEL PRO v8.8.4 IN ---")
  (setq ss (ssget "X" '((0 . "INSERT") (410 . "Model"))))
  (if (not ss) (progn (alert "No hay bloques.") (exit)))
  
  (setq bl '() i 0 q (chr 34) c40 (chr 40) c41 (chr 41))
  (repeat (sslength ss) (setq e (ssname ss i) o (vlax-ename->vla-object e) n (vl-catch-all-apply 'vla-get-EffectiveName (list o))) (if (vl-catch-all-error-p n) (setq n (vla-get-Name o))) (setq n (vl-princ-to-string n)) (if (not (assoc n bl)) (setq bl (cons (cons n 0) bl))) (setq i (1+ i)))
  
  (setq current bl df (strcat (if (getvar "TEMPPREFIX") (getvar "TEMPPREFIX") "C:/Temp/") "ex_in.dcl") f (open df "w"))
  (write-line (strcat "ex : dialog { label=" q "Excel Pro IN v8.5" q "; :column { :edit_box { key=" q "f" q "; label=" q "Filtro:" q "; } :list_box { key=" q "l" q "; multiple_select=true; height=20; width=55; } } :row { :button { key=" q "accept" q "; label=" q "GENERAR" q "; is_default=true; } :button { key=" q "cancel" q "; label=" q "Cerrar" q "; is_cancel=true; } } }") f) (close f)
  
  (defun ULV () (start_list "l") (foreach x current (add_list (strcat (if (= (cdr x) 1) "[X] " "[ ] ") (car x)))) (end_list))
  (defun UF (v / s) (setq s (strcase v) current '()) (foreach x bl (if (wcmatch (strcase (car x)) (strcat "*" s "*")) (setq current (append current (list x))))) (ULV))
  (defun TS (v / idx bn os ns cm it) (setq cm (strcat c40 v c41) idx (read cm)) (foreach i idx (setq bn (car (nth i current)) it (assoc bn bl) os (cdr it) ns (if (= os 1) 0 1) bl (subst (cons bn ns) it bl) current (subst (cons bn ns) (nth i current) current))) (ULV))
  
  (setq id (load_dialog df))
  (if (and id (new_dialog "ex" id))
    (progn (ULV) (action_tile "f" "(UF $value)") (action_tile "l" "(TS $value)") (action_tile "accept" "(done_dialog 1)") (action_tile "cancel" "(done_dialog 0)") (setq rs (start_dialog)) (unload_dialog id)
      (if (= rs 1)
        (progn (setq sb '()) (foreach x bl (if (= (cdr x) 1) (setq sb (cons (car x) sb))))
          (if sb 
            (progn (setq mi 53 t_s '("BLOQUE") ty '() i 0 hm nil ft 0.0 fc 0.0)
              (repeat (sslength ss) (setq e (ssname ss i) o (vlax-ename->vla-object e) n (vl-princ-to-string (vl-catch-all-apply 'vla-get-EffectiveName (list o)))) (if (and (member n sb) (not (member n ty))) (setq ty (append ty (list n)))) (setq i (1+ i)))
              (foreach b_n ty (setq b_o (EX_GBO b_n)) (foreach t_g b_o (if (not (member t_g t_s)) (setq t_s (append t_s (list t_g))))))
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
                        (if (wcmatch (strcase n) "*CONDULETA*")
                            (progn (vlax-put-property (vlax-get-property xs 'Range (strcat (EX_GEX 55) (itoa r))) 'Value2 pz) (setq fc (+ fc pz)))
                        )
                    )
                    (setq r (1+ r))
                  )
                )
                (setq i (1+ i))
              )
              (vl-catch-all-apply 'vlax-invoke (list (vlax-get-property (vlax-get-property xs 'Range "A1") 'CurrentRegion) 'AutoFilter))
              (vl-catch-all-apply 'vlax-invoke (list (vlax-get-property xs 'UsedRange) 'Sort (vlax-get-property xs 'Range "A2") 1))
              (setq cr (+ r 2))
              (if (> ft 0.0) (setq cr (EX_AnexoTEE xs cr (strcat (EX_GEX 54) "2:" (EX_GEX 54) (itoa (1- r))))))
              (if (> fc 0.0) (setq cr (EX_AnexoCond xs cr (strcat (EX_GEX 55) "2:" (EX_GEX 55) (itoa (1- r))))))
              (if hm (progn (setq range_str (strcat (EX_GEX mi) "2:" (EX_GEX mi) (itoa (1- r)))) (setq cr (EX_AnexoPipe xs cr range_str mi))))
              
              ;; Forzar actualizacion y expansion total
              (vlax-put-property xa 'ScreenUpdating :vlax-true)
              (vlax-invoke-method (vlax-get-property (vlax-get-property xs 'Cells) 'Columns) 'AutoFit)
              
              ;; Ocultar columnas de calculo con proteccion de error
              (vl-catch-all-apply 'vlax-put-property 
                (list (vlax-get-property xs 'Range (strcat (EX_GEX mi) ":" (EX_GEX 55))) 'Hidden :vlax-true))
              
              (princ "\n>>> BOM GENERADA CON EXPANSIÓN TOTAL (v9.6.0) <<<")
            )
          )
        )
      )
    )
  )
  (if (and df (vl-file-size df)) (vl-file-delete df)) (princ)
)

(princ "\n--- EXCEL PRO v9.6.0 [IMPERIAL PRO] ---") (princ)
