;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:cl-user)


(defpackage #:splt
  (:use #:cl #:utils #:ccom4 #:ccoffice #:achar #:wax)
  (:export #:categories
           #:newfile-name
           #:find-end
           #:1region
           #:regions
           #:noncategorical-rows
           #:prune
           #:split-autofilter
           #:split-chunks
           #:wbooks->wsheets
           #:*copy-paste-wait*
           #:copy-paste
           #:find-last-row
           #:copy-ws
           #:wsheets->1wsheet
           #:join-worksheets
           ))
