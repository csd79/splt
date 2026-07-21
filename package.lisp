;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:cl-user)


(defpackage #:splt
  (:use #:cl #:ccom4 #:ccoffice #:achar)
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
           ;#:splt-start
           )
