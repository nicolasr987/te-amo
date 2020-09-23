call fala
sub fala()
set audio =createobject("SAPI.SPVOICE")
audio.volume=100
audio.rate = -1
call inicio
end sub
dim i,tipo,n,audio
call inicio()
sub inicio() 
tipo=inputbox("escreva o nome do amor da sua vida")
audio.speak ("a o amor dasua vida é: "& tipo & + vbnewline & _
              "te  amo: "& tipo &"")
end sub