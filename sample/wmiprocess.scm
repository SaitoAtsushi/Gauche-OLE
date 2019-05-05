#!/usr/bin/env gosh
(use win.ole)
(use gauche.collection)

(define oLocator (make-ole "WbemScripting.SWbemLocator"))
(define oService (oLocator 'ConnectServer))
(define oClassSet (oService 'ExecQuery "Select * From Win32_Process"))

(for-each (lambda(x)
            (display (~ x 'Description))
            (display ":")
            (display (~ x 'ProcessId))
            (newline))
          oClassSet)

(ole-release!)
