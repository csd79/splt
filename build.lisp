;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 

(in-package "CL-USER")

(load "c:\\Users\\cselovszkid\\.lispworks")

(asdf:load-system "splt")

(lw:deliver 'splt:run
    "c:\\Users\\cselovszkid\\common-lisp\\splt\\splt.exe"
    5
;    :console :io
    :multiprocessing t
    :console t
    :error-on-interpreted-functions t
    :keep-eval t
    :keep-lisp-reader t
    :symbol-names-action nil
    :startup-bitmap-file nil
    :KILL-DSPEC-TABLE NIL
    :keep-package-manipulation t


    :keep-debug-mode t



;    :compact t
;    :keep-clos-object-printing t
;    :keep-function-name t
;    :keep-load-function t
;    :keep-structure-info t
;    :keep-top-level t
;    :make-instance-keyword-check t
;    :no-symbol-function-usage t
;    :packages-to-keep-symbol-names t


    )
