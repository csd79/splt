(require "automation")

(defsystem "splt"
  :description "Splitting the atom"
  :author      "Denes Cselovszki <denes.cselovszki@gmail.com>"
  :version     "0.07"
  :depends-on  ("ccom4" "ccoffice" "cl-ppcre" "achar")
  :serial      t
  :components  ((:file "package")
                (:file "fli-templates")
                (:file "splt")))
