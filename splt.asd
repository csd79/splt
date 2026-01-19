(require "automation")

(defsystem "splt"
  :description "Splitting the atom"
  :author      "Denes Cselovszki <denes.cselovszki@gmail.com>"
  :version     "0.06"
  :depends-on  ("ccom3" "msoffice" "cl-ppcre" "achar")
  :serial      t
  :components  ((:file "package")
                (:file "fli-templates")
                (:file "splt")))
