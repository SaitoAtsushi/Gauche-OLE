#!/usr/bin/env gosh
(use win.ole)
(use gauche.collection)

(define IE (make-ole "InternetExplorer.Application"))
(set! (~ IE 'Visible) #t)
(while (~ IE 'busy) (sys-sleep 1))
(IE 'Navigate "http://example.com/")
(while (~ IE 'busy) (sys-sleep 1))
(display
 (? IE 'Document ! 'getElementsByTagName "div" ! 'item 0 ? 'innerText))

(ole-release!)
