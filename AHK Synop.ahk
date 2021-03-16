f9::
wbk := ComObjGet("D:\Users\Downloads\31.xls")
wbc := wbk.Sheets("Input")
Send {Tab}

InputBox, jam, Jam UTC, Masukkan Jam UTC., ,200,100
if ErrorLevel{
    MsgBox, Tombol CANCEL ditekan.
    Esc::ExitApp
}
else{
    MsgBox, Input Metar jam %jam% UTC
}
if(jam<6)
    space:= 5*jam
else if(jam>5 and jam<12)
    space:= 14 + (5*jam)
else if(jam>11 and jam<18)
    space:= 28 + (5*jam)
else if(jam>17 and jam<24)
    space:= 42 + (5*jam)

dataAngin:= 10+space
isi:=wbc.Range("B"dataAngin).Value
if(isi=3){
   Send {Tab} 
}
if(isi=4){
   Send {Down 2} 
   Send {Enter}
   Send {Tab}
}

arahAngin:= 12+space
isi:= wbc.Range("B"arahAngin).Value
SendInput, % Floor(isi)
Send {Tab}

kecAngin:= 13+space
isi:= wbc.Range("B"kecAngin).Value
SendInput, % Floor(isi)
Send {Tab}

vv:= 10+space
isi:= wbc.Range("D"vv).Value
SendInput, % Floor(isi)
Send {Tab}

dataCuaca:= 10+space
isi:= wbc.Range("I"dataCuaca).Value
if(isi>1){
    Loop, % isi
    {
        Send {Down} 
    }
    Send {Enter}
    Send {Tab}
}
else
    Send {Tab}

;sekip     
Send {Tab 3}

degPanas:=10+space
isi:= wbc.Range("J"degPanas).Value
SendInput, % isi
Send {Tab}

pDibaca:=11+space
isi:= wbc.Range("J"pDibaca).Value
SendInput, % isi
Send {Tab}

qff:=11+space
isi:= wbc.Range("K"qff).Value
SendInput, % isi
Send {Tab}

qfe:=11+space
isi:= wbc.Range("L"qfe).Value
SendInput, % isi
Send {Tab}

bolaKering:=10+space
isi:= wbc.Range("M"bolaKering).Value
SendInput, % isi
Send {Tab}

bolaBasah:=12+space
isi:= wbc.Range("M"bolaBasah).Value
SendInput, % isi
Send {Tab 3}

;sekip awan rendah
Send {Tab}

nAwanMenengah:=10+space
isi:= wbc.Range("Z"nAwanMenengah).Value
if(isi>0){
    Send {Down}
    Loop, % isi
    {
        Send {Down} 
    }
    Send {Enter}
    Send {Tab}
}
else
    Send {Tab}

nAwanTinggi:=12+space
isi:= wbc.Range("Z"nAwanTinggi).Value
if(isi>0){
    Send {Down}
    Loop, % isi
    {
        Send {Down} 
    }
    Send {Enter}
    Send {Tab}
}
else
    Send {Tab}

;sekip awan menengah jenis dan tinggi jenis, arah derajat
Send {Tab 3}

tinggiAwan1:=11+space
isi:= wbc.Range("AE"tinggiAwan1).Value
SendInput, % Floor(isi)
Send {Tab}

;sekip arah derajat 2
Send {Tab}

tinggiAwan2:=13+space
isi:= wbc.Range("AE"tinggiAwan2).Value
SendInput, % Floor(isi)
Send {Tab}

tertutupAwan:=10+space
isi:= wbc.Range("AG"tertutupAwan).Value
if(isi>0){
    Send {Down}
    Loop, % isi
    {
        Send {Down} 
    }
    Send {Enter}
    Send {Tab}
}
else
    Send {Tab}

;sekip jenis CL1-mari
Send {Tab 12}

keadaanTanah:=10+space
isi:= wbc.Range("AV"keadaanTanah).Value
if(isi>0){
    Send {Down}
    Loop, % isi
    {
        Send {Down} 
    }
    Send {Enter}
    Send {Tab}
}
else
    Send {Tab}

catatanTanah:=12+space
isi:= wbc.Range("AV"catatanTanah).Value
SendInput, % isi
MsgBox Pilih nama observer secara manual

return