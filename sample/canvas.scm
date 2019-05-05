#!/usr/bin/env gosh
(use win.ole)

(define IE (make-ole "InternetExplorer.Application"))
(set! (~ IE 'Visible) #t)
(while (~ IE 'busy) (sys-sleep 1))
(IE 'Navigate "about:blank")
(define canvas (? IE 'Document ! 'createElement "canvas"))
(set! (~ canvas 'height) 300)
(set! (~ canvas 'width) 400)

(? IE 'Document ? 'body ! 'appendChild canvas)

(define context (canvas 'getContext "2d"))
(context 'beginPath)
(context 'moveTo 30 10)
(context 'lineTo 45 50)
(context 'lineTo 5 50)
(set! (~ context 'lineWidth) 5)
(set! (~ context 'strokeStyle) "black")
(context 'closePath)
(context 'stroke)
(set! (~ context 'fillStyle) "red")

(context 'fillRect 20 40 50 100)
(context 'arc 70 45 30 0 6 #t)
(context 'fill)
(ole-release!)
