(in-package :splt)


;;; ---------------------------------------------------------------
;;; Operations


(defun split-autofilter (file column &optional (start 2))
  "Split first worksheet of an Excel workbook on given COLUMN using the autofilter."
  (with-property-accessors
    (setf (property-accessors-on) t)
    (with-workbook (:wbook wbook :open file :wsvars (ws1) :close t)
      (let* ((last-row (last-row ws1))
             (raw      (?value2 (range ws1 column start column last-row)))
             (list     (loop for i from 0 below (array-dimension raw 0) collecting (aref raw i 0)))
             (categories (remove-duplicates list :test #'string=)))
        (dolist (categorie categories)
          (let* ((new-ws   (xselect> ws1 `((,column ,categorie)) :paste-all t :paste-column-widths t
                                     :header-row (1- start)))
                 (new-wb   (?parent new-ws))
                 (new-file (make-pathname
                            :name (concatenate 'string categorie " - " (pathname-name file))
                            :type (pathname-type file)
                            :defaults file)))
            (print new-file)
            (!saveas new-wb (namestring new-file))
            (!close new-wb)))))
    (print "OK")))


(defun split-chunks (file col &optional (start 2))
  "Split first worksheet of an Excel workbook on given column by copying source file and deleting unnecessary rows."
  (with-property-accessors
    (setf (property-accessors-on) t)
    (let ((categories '()))
      (with-workbook (:wbook wbook :open file :wsvars (ws1) :save nil :close t)
        (multiple-value-bind (last right) (last-row ws1)
          (let* ((raw  (?value2 (range ws1 col start col last)))
                 (list (loop for i from 0 below (array-dimension raw 0) collecting (aref raw i 0))))
            (setf categories (remove-duplicates list :test #'string=))))
        (dolist (categorie categories)
          (with-workbook (:wbook wbook :open file :wsvars (ws1) :save nil :close nil)
            (let* ((last (last-row ws1))
                   (new-file (make-pathname
                              :name (concatenate 'string categorie " - " (pathname-name file))
                              :type (pathname-type file)
                              :defaults file)))
              (!saveas wbook (namestring new-file))
              (loop for r from last downto start doing
                    (when (string/= (xcell ws1 col r) categorie)
                      (!delete (?entirerow (range ws1 1 r right r)))))
              (!save wbook)
              (print new-file)
              (!close wbook))))))))


(defun wbooks->wsheets (directory &optional (rowheight nil))
  "Collect the first worksheet of every workbook in DIRECTORY into a new workbook."
  (with-property-accessors
    (setf (property-accessors-on) t)
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
                (setf (?rowheight (used-range (!item outwsheets (1- (?count outwsheets))))) rowheight)))))))))


#|(defun run ()
  "Run an operation from the command line."
  (let ((file (second sys:*line-arguments-list*)))
    (when (probe-file file)
      (format t "Oszlop: ")
      (let ((title (read-line)))
        (split-autofilter file title)
        (read-line))))))|#


;;; ---------------------------------------------------------------
;;; Sandbox


(defparameter *nd*
  "c:\\Users\\cselovszkid\\Downloads\\2026.01.04. Januári béremelések\\3 Ped PedNOKS\\Közös 0008\\v3\\")


(defun test ()
  (wbooks->wsheets *dir* 124.5))
