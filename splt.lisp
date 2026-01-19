(in-package :splt)


;;; ---------------------------------------------------------------
;;; Utilities


(defun categories (ws column &optional (start 2))
  "List unique strings in COLUMN of the first worksheet of FILE."
  (multiple-value-bind (lowest rightmost)
      (last-row ws)
    (let* ((raw  (?value2 (range ws column start column lowest)))
           (cats (remove-duplicates
                  (column->row raw :element-type 'string)
                  :test #'string=)))
      (values cats lowest rightmost))))

(defun newfile-name (category oldfile &optional (control-string "~a - ~a"))
  "Construct filename of the resulting file(s)."
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
  "List the number of rows that don't belong to CATEGORY."
  (let ((rows '())
        (raw  (?value2 (range wsheet column first column (last-row wsheet)))))
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
      (!delete (?entirerow (range wsheet 1 (first region) 1 (second region)))))))


;;; ---------------------------------------------------------------
;;; Operations


(defun split-autofilter (file column &optional (start 2))
  "Split first worksheet of an Excel workbook on given COLUMN using the autofilter."
  (with-workbook (:wbook wbook :open file :wsvars (ws1) :close t)
    (let ((categories (categories ws1 column start)))
      (loop for category across categories doing
        (let* ((new-ws   (xselect> ws1 `((,column ,category)) :paste-all t :paste-column-widths t
                                   :header-row (1- start)))
               (new-wb   (?parent new-ws))
               (new-file (newfile-name category file)))
          (!saveas new-wb (namestring new-file))
          (print new-file)
          (!close new-wb))))))


(defun split-chunks (file column &optional (start 2))
  "Split first worksheet of an Excel workbook on given column by copying source file and deleting unnecessary rows."
  (with-workbook (:wbook wbook :open file :wsvars (ws1) :save nil :close t)
    (let ((categories (categories ws1 column start)))
      (pprint categories)
      (loop for category across categories doing
            (with-workbook (:wbook wbook :open file :wsvars (nws1) :save nil :close nil)
        (let ((newfile-name (newfile-name category file)))
          (!saveas wbook (namestring newfile-name))
          (prune nws1 column category start :verbose t)
          (!save wbook)
          (print newfile-name)
          (!close wbook)))))))


(defun wbooks->wsheets (directory &optional (rowheight nil))
  "Collect the first worksheet of every workbook in DIRECTORY into a new workbook."
    (cclet* ((files  (sort (uiop:directory-files *dir*) #'astring< :key #'pathname-name))
             (xlsxs  (remove-if #'(lambda (pathname)
                                    (stringnot-equal (pathname-type pathname) "xlsx"))
                                files))
             (result (concatenate 'string directory "_EREDMÉNY_.xlsx"))
             (excel  (com:create-object :progid "Excel.Application")))
      (with-workbook (:app excel :wbook outbook :wsvars (outws) :save t :close t)
        (!saveas outbook result)
        (cclet* ((outwsheets (?worksheets outbook)))
          (dolist (infile xlsxs)
            (cclet* ((last-sheet (!item outwsheets (?count outwsheets)))
                     (inname     (pathname-name infile)))
              (with-workbook (:app excel :wbook inbook :open (namestring infile) :wsvars (inws) :close t)
                (!copy inws last-sheet))
              (print inname)
              (setf (?name (!item outwsheets (1- (?count outwsheets)))) inname)
              (when rowheight
                (setf (?rowheight (used-range (!item outwsheets (1- (?count outwsheets))))) rowheight))))))))


(defmacro runenv ((args-var) &body body)
  "Put command line arguments into ARGS-VAR, then execute BODY with MS Office property accessors on."
  `(let ((,args-var (rest sys:*line-arguments-list*)))
     (with-property-accessors
       (setf (property-accessors-on) t)
       ,@body)))
     


;;; ---------------------------------------------------------------
;;; Sandbox


(defparameter *dir*
;  "c:\\Users\\cselovszkid\\Downloads\\2026.01.04. Januári béremelések\\3 Ped PedNOKS\\Közös 0008\\v3\\"
  "c:\\Users\\cselovszkid\\Downloads\\2026.01.04. Januári béremelések\\3 Ped PedNOKS\\Közös 0008\\v4\\"
  )

(defparameter *file*
  "c:\\Users\\cselovszkid\\Downloads\\2026.01.04. Januári béremelések\\3 Ped PedNOKS\\Közös 0008\\vx\\IT0008_illetmények.xlsx"
  )


(defun test ()
  (runenv (args)
;    (split-autofilter *file* 21 4)
    (split-chunks *file* 21 4)
;    (wbooks->wsheets *dir* 124.5)
  ))
