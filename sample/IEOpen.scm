#!/usr/bin/env gosh
(use win.ole)

(define IE (make-ole "InternetExplorer.Application"))
(set! (~ IE 'Visible) #t)
(while (~ IE 'busy) (sys-sleep 1))
(IE 'Navigate "http://www.yahoo.co.jp")
(ole-release!)
