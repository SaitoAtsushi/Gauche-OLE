#!/usr/bin/env gosh
(use win.ole)

(define request (make-ole "Microsoft.XMLHTTP"))

(define (checkstatus)
  (when (and (= 4 (~ request 'readystate))
             (= 200 (~ request 'status)))
    (display (~ request 'responseText))
    (ole-release!)))

(set! (~ request 'onreadystatechange) checkstatus)
(request 'open "GET" "http://www.example.com/" #t)
(request 'send '())
