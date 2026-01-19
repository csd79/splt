(require "automation")

(defsystem "splt"
  :description "Splitting the atom"
  :author      "Denes Cselovszki <denes.cselovszki@gmail.com>"
  :version     "0.03"
  :depends-on  ("ccom3" "msoffice" "cl-ppcre" "achar")
  :serial      t
  :components  ((:file "package")
                (:file "fli-templates")
                (:file "splt")))
