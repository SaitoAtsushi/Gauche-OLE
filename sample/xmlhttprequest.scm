#!/usr/bin/env gosh
(use win.ole)

(define request (make-ole "Microsoft.XMLHTTP"))
(request 'open "GET" "http://www.example.com/" #f)
(request 'send '())
(when (equal? 200 (~ request 'status))
  (display (? request 'responseText)))
(ole-release!)
