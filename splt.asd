(require "automation")

(defsystem "splt"
  :description "Splitting the atom"
  :author      "Denes Cselovszki <denes.cselovszki@gmail.com>"
  :version     "0.08"
  :depends-on  ("utils" "ccom4" "ccoffice" "cl-ppcre" "achar" "wax")
  :serial      t
  :components  ((:file "package")
                (:file "fli-templates")
                (:file "splt")))
