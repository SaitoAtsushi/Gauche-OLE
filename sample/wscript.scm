#!/usr/bin/env gosh
(use win.ole)

(define wsh (make-ole "WScript.Shell"))
(define exec (wsh 'Exec "cmd /C ver"))
(until (~ exec 'StdOut 'AtEndOfStream)
  (display (? exec 'StdOut ! 'ReadLine)))
(ole-release!)
