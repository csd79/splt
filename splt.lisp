(in-package :splt)
#.(enable-ccom-syntax)


;;; ---------------------------------------------------------------
;;; Utilities


(defun categories (ws column &optional (start 2))
  "List unique strings in COLUMN of worksheet WS."
  (multiple-value-bind (lowest rightmost)
      (last-row ws)
    (let* ((raw  (?'value2 (range ws column start column lowest)))
           (cats (remove-duplicates
                  (column->row raw :element-type 'string)
                  :test #'string=)))
      (values cats lowest rightmost))))

(defun newfile-name (category oldfile &optional (control-string "~a - ~a"))
  "Construct name of output file for CATEGORY."
  (make-pathname
   :name (format nil control-string category (pathname-name oldfile))
   :type (pathname-type oldfile)
   :defaults oldfile))


;;; ---------------------------------------------------------------
;;; Chunk pruning


(defun find-end (seq &key (start 0) (diff 1))
  "Find the end of arithmetic sequence of difference = 1 starting at START."
  (let ((prev (svref seq start))
        (last (1+ start)))
    (loop for i from last below (length seq)
          for e = (svref seq i)
          while (= (- e prev) diff)
          do (setf prev e last i)
          finally (return (if (= i last)
                            start last)))))

(defun 1region (seq &optional (regions '()) (start 0))
  "Recursively build up a list of regions (sublists denoting the start and end of a chunk)."
  (if (>= start (length seq))
    regions
    (let ((next (find-end seq :start start)))
      (1region seq
               (cons (list (svref seq start) (svref seq next)) regions)
               (1+ next)))))

(defun regions (seq)
  "Start the recursive sequence of 1REGION."
  (sort (1region seq) #'> :key #'first))

(defun noncategorical-rows (wsheet column category first)
  "List the rows (by row number) that don't belong to CATEGORY."
  (let ((rows '())
        (raw  (?'value2 (range wsheet column first column (last-row wsheet)))))
    (loop for e across (column->row raw :element-type 'string)
          for i from 0 doing
          (unless (string= e category)
            (push (+ i first) rows)))
    rows))

(defun prune (wsheet column category start &key (verbose nil))
  "Delete noncategorical regions."
  (let* ((rows    (sort (noncategorical-rows wsheet column category start) #'<))
         (regions (regions (coerce rows 'vector))))
    (when verbose (format t "~a   " wsheet))
    (dolist (region regions)
      (when verbose (format t "~a   " region))
      (!'delete (?'entirerow (range wsheet 1 (first region) 1 (second region)))))))


;;; ---------------------------------------------------------------
;;; Splitting


(defun split-autofilter (file column &optional (start 2))
  "Split first worksheet of an Excel workbook on given COLUMN using the autofilter."
  (with-ccom-context
    (with-workbook (:wbook wbook :open file :wsvars (ws1) :close t)
      (let ((categories (categories ws1 column start)))
        (loop for category across categories doing
              (let* ((new-ws   (xselect> ws1 `((,column ,category)) :paste-all t :paste-column-widths t
                                         :header-row (1- start)))
                     (new-wb   (?'parent new-ws))
                     (new-file (newfile-name category file)))
                (!'saveas new-wb (namestring new-file))
                (print new-file)
                (!'close new-wb)))))))


(defun split-chunks (file column &optional (start 2))
  "Split first worksheet of an Excel workbook on given column by copying source file and deleting unnecessary rows."
  (with-ccom-context
    (with-workbook (:wbook wbook :open file :wsvars (ws1) :save nil :close t)
      (let ((categories (categories ws1 column start)))
        (pprint categories)
        (loop for category across categories doing
              (with-workbook (:wbook wbook :open file :wsvars (nws1) :save nil :close nil)
                (let ((newfile-name (newfile-name category file)))
                  (!'saveas wbook (namestring newfile-name))
                  (prune nws1 column category start :verbose t)
                  (!'save wbook)
                  (print newfile-name)
                  (!'close wbook))))))))
  

(defun wbooks->wsheets (directory &optional (rowheight nil))
  "Collect the first worksheet of every workbook in DIRECTORY into a new workbook."
  (with-ccom-context
    (cclet* ((files  (sort (uiop:directory-files *dir*) #'astring< :key #'pathname-name))
             (xlsxs  (remove-if #'(lambda (pathname)
                                    (string-not-equal (pathname-type pathname) "xlsx"))
                                files))
             (result (concatenate 'string directory "_EREDMÉNY_.xlsx"))
             (excel  (com:create-object :progid "Excel.Application")))
      (with-workbook (:app excel :wbook outbook :wsvars (outws) :save t :close t)
        (!'saveas outbook result)
        (cclet* ((outwsheets (?'worksheets outbook)))
          (dolist (infile xlsxs)
            (cclet* ((last-sheet (!'item outwsheets (?'count outwsheets)))
                     (inname     (pathname-name infile)))
              (with-workbook (:app excel :wbook inbook :open (namestring infile) :wsvars (inws) :close t)
                (!'copy inws last-sheet))
              (print inname)
              (setf (?'name (!'item outwsheets (1- (?'count outwsheets)))) inname)
              (when rowheight
                (setf (?'rowheight (used-range (!'item outwsheets (1- (?'count outwsheets))))) rowheight)))))))))


;;; ---------------------------------------------------------------
;;; Joining


(defparameter *copy-paste-wait* 0.1)

(defun copy-paste (source-range target-wsheet target-range &key (colwidth nil))
  "Copy contents of SRC-RANGE into TRG-RANGE of TRG-WSHEET."
  (!'copy source-range)
  (sleep *copy-paste-wait*)
  (!'paste target-wsheet target-range)
  (sleep *copy-paste-wait*)
  (when colwidth
    (!'pastespecial target-range 8)
    (sleep *copy-paste-wait*))
  )


(defun find-last-row (worksheet start column)
  "Find the last row of WORLSHEET (starting from START), where COLUMN is not empty."
  (let ((xl-last (last-row worksheet)))
    (loop for r from xl-last downto start doing
          (when (not (empty-cell-p (xcell worksheet column r)))
            (return r))
          finally (return (1- start)))))


(defun copy-ws (source-ws target-ws starting-row check-column width &key (colwidth nil)); &optional (message-stream nil))
  "Copy contents of SOURCE-WS (columns 1 through WIDTH) from STARTING-ROW to the last row where CHECK-COLUMN is not empty, and paste it into TARGET-WS starting beneath the last row where CHECK-COLUMN is not empty."
  (cclet* ((src-last     (find-last-row source-ws starting-row check-column))
           (trg-starting (1+ (find-last-row target-ws starting-row check-column)))
           (trg-last     (- (+ src-last trg-starting) starting-row))
           (src-range    (range source-ws 1 starting-row width src-last))
           (trg-range    (range target-ws 1 trg-starting width trg-last))
           (row-count    (- src-last starting-row -1)))
#|    (when message-stream
      (format message-stream "~a-~a   -->   ~a-~a~%" starting-row src-last trg-starting trg-last))|#
    (unless (zerop row-count)
      (copy-paste src-range target-ws trg-range :colwidth colwidth))
    row-count))


(defun wsheets->1sheet (infile outfile starting-row check-column width)
  "Copy the contents of all worksheets from INFILE into one worksheet in OUTFILE (a new file)."
  (with-ccom-context
    (cclet* ((excel (com:create-object :progid "Excel.Application")))
      (with-workbook (:app excel :wbook outbook :wsvars (outsheet) :save t :close t)
        (!'saveas outbook outfile)
        (with-workbook (:app excel :wbook inbook :open (namestring infile)  :wsvars (first) :save nil :close t)
          (!'copy first outsheet)
          (cclet* ((insheets (?'worksheets inbook))
                   (count    (?'count insheets))
                   (outsheet (!'item (?'worksheets outbook) 1)))
            (loop for i from 2 upto count
                  for insheet = (!'item insheets i) doing
                  (copy-ws insheet outsheet starting-row check-column width))))))))


(defun join-worksheets (infiles outfile obj &key (ws-number 1) (starting-row 2) (check-column 3) (width 30))
;                                (message-fn nil) (step-fn nil))
  "Copy contents of WSHEET from all INFILES into one sheet in OUTFILE (a new file)."
#|  (flet ((message (&rest args) (when message-fn (apply message-fn args)))
         (step (&rest args) (when step-fn (apply step-fn args))))|#
    (with-ccom-context
      (cclet* ((excel  (com:create-object :progid "Excel.Application"))
               (first? t)
               (copied 0))
        (excellerate (excel)
          ;; Open output file
          (with-local-copies ((outf/ outfile :new-file t))
            (with-workbook (:app excel :wbook outbook :wsvars (outsheet) :save t :close t)
              (!'saveas outbook outf/)
              ;; Loop over input files.
              (loop for infile in infiles doing
                    ;; If infile exists:
                    (if (probe-file infile)
                      ;; Open infile for editing. In case of a condition, sign unhandled error on display.
                      (multiple-value-bind (result condition)
                          (ignore-errors
                            (with-local-copies ((inf/ infile))
                              (with-workbook (:app excel :wbook inbook :open inf/ :wsvars (1st-in)
                                              :save t :close t)
                                ;; When input file exists, find the required worksheet in it.
                                (cclet* ((insheet (!'item (?'worksheets inbook) ws-number)))
                                  ;; If this isqrt the first input file, copy header as well.
                                  (if first?
                                    (progn
                                      (setf copied (copy-ws insheet outsheet 1 check-column width :colwidth t))
                                      (setf first? nil))
                                    ;; Otherwise, just copy the body.
                                    (setf copied (copy-ws insheet outsheet starting-row check-column width)))))))
                        (declare (ignore result))
                        (if (not condition)
                          ;; Print message about whether the file was processed or not.
                          (disp obj "~a.xlsx: ~a sor másolva.~2%" (pathname-name infile) copied)
                          ;; ERROR
                          (let ((error-message (format nil "Kezeletlen hiba a ~a.xlsx fájl használatakor!~2%"
                                                       (pathname-name infile))))
                            (disp obj error-message)
                            (wax::queue-errorlog obj error-message))
                            ))
                      ;; If infile doesn't exist, display warning.
                      (disp obj "~a.xlsx nem található.~2%" (pathname-name infile)))
                    ;; Step progress bar.
                    (pstep obj))
              ))))))


;;; ---------------------------------------------------------------
;;; Sandbox


#|(defmacro runenv ((args-var) &body body)
  "Put command line arguments into ARGS-VAR, then execute BODY with MS Office property accessors on."
  `(let ((,args-var (rest sys:*line-arguments-list*)))
     ,@body))|#





#|(defparameter *dir*
;  "c:\\Users\\cselovszkid\\Downloads\\2026.01.04. Januári béremelések\\3 Ped PedNOKS\\Közös 0008\\v3\\"
  "c:\\Users\\cselovszkid\\Downloads\\2026.01.04. Januári béremelések\\3 Ped PedNOKS\\Közös 0008\\v4\\"
  )

(defparameter *file*
  "c:\\Users\\cselovszkid\\Downloads\\2026.01.04. Januári béremelések\\3 Ped PedNOKS\\Közös 0008\\vx\\IT0008_illetmények.xlsx"
  )|#

#|(defparameter *file*
  "C:\\Users\\cselovszkid\\Downloads\\2026.06.01. Mariann szétszedés\\1 PSZE\\KK_PSZE_2026_Június_BM_11465_átadandó elismerések.xlsx")|#

#|(defparameter *file*
  "C:\\Users\\cselovszkid\\Downloads\\2026.06.01. Mariann szétszedés\\2 MEO\\MEO_2025__Pedagógusnap 120 fõ_ÁT_BM_885_közös ABC.xlsx")|#


#|(defun hh ()
  (with-ccom-context
    (with-workbook (:wbook outbook :open *mo* :wsvars (outsheet) :save nil :close t)
      (print (last-row outsheet)))))
    

(defparameter *mi* "C:\\Users\\cselovszkid\\Downloads\\--------------------\\Táblázatok\\Döntési tábla_2026. június 23. igazgatói pályázat_5 év_SZÉTBONTOTT (002).xlsx")
(defparameter *mo* "C:\\Users\\cselovszkid\\Downloads\\--------------------\\Táblázatok\\grr Döntési tábla_2026. június 23. igazgatói pályázat_5 év.xlsx")|#


#|(defun test ()
  (runenv (args)
    (with-ccom-context        
;     (split-autofilter *file* 21 4)
     (split-chunks *file* 4 6)
;     (wbooks->wsheets *dir* 124.5)
     )))|#






#.(disable-ccom-syntax)
