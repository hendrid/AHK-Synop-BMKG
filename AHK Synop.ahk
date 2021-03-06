^!s::
InputBox, nama, Tanggal, Masukkan Tanggal dgn format dd/mm/yy `nContoh: 1/1/21,,200,160
StringSplit, tanggal_array, nama, '/'
format=xls
if(tanggal_array2=1)
    tanggal_array2=1.Januari
else if(tanggal_array2=2)
    tanggal_array2=2.Pebruari
else if(tanggal_array2=3)
    tanggal_array2=3.Maret
else if(tanggal_array2=4)
    tanggal_array2=4.April
else if(tanggal_array2=5)
    tanggal_array2=5.Mei
else if(tanggal_array2=6)
    tanggal_array2=6.Juni
else if(tanggal_array2=7)
    tanggal_array2=7.Juli
else if(tanggal_array2=8)
    tanggal_array2=8.Agustus
else if(tanggal_array2=9)
    tanggal_array2=9.September
else if(tanggal_array2=10)
    tanggal_array2=10.Oktober
else if(tanggal_array2=11)
    tanggal_array2=11.Nopember
else if(tanggal_array2=12)
    tanggal_array2=12.Desember

namafile=D:\SYNOP_20%tanggal_array3%\%tanggal_array2%\%tanggal_array1%.%format%
wbk := ComObjGet(namafile)
wbc := wbk.Sheets("Input")
Send {Tab}

InputBox, jam, Jam UTC, Masukkan Jam UTC,,200,130
if ErrorLevel{
    MsgBox, Tombol CANCEL ditekan.
    Esc::ExitApp
}
else{
    MsgBox, Input Synop jam %jam% UTC
}
if(jam<6)
    space:= 5*jam
else if(jam>5 and jam<12)
    space:= 14 + (5*jam)
else if(jam>11 and jam<18)
    space:= 28 + (5*jam)
else if(jam>17 and jam<24)
    space:= 42 + (5*jam)

SetFormat,float,0.1

dataAngin:= 10+space
isi:=wbc.Range("B"dataAngin).Value
if(isi=3){
    Send {Down} 
    Send {Enter}
    Send {Tab} 
}
if(isi=4){
   Send {Down 2} 
   Send {Enter}
   Send {Tab}
}

arahAngin:= 12+space
isi:= wbc.Range("B"arahAngin).Value
if(isi="calm"){
    isi=0
}
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
if(isi>0){
    Loop, % isi
    {
        Send {Down} 
    }
    Send {Enter}
    Send {Tab}
}
else
    Send {Tab}

ww:=11+space
isi:= wbc.Range("I"ww).Value
if(isi != ""){
    if(isi="cld dev unk")
        isi=0
    else if (isi="cld decr")
        isi=1
    else if (isi="cld unch")
        isi=2
    else if (isi="cld incr")
        isi=3
    else if (isi="smoke")
        isi=4
    else if (isi="haze")
        isi=5
    else if (isi="dust 06")
        isi=6
    else if (isi="dust 07")
        isi=7
    else if (isi="sand 07")
        isi=7
    else if (isi="dw")
        isi=8
    else if (isi="dust whirl")
        isi=8
    else if (isi="ss")
        isi=8
    else if (isi="sand storm")
        isi=8
    else if (isi="dw 09")
        isi=9
    else if (isi="dust whirl 09")
        isi=9
    else if (isi="ss 09")
        isi=9
    else if (isi="sand storm 09")
        isi=9
    else if (isi="mist")
        isi=10
    else if (isi="shallow fog 11")
        isi=11
    else if (isi="shallow fog 12")
        isi=12
    else if (isi="lightning")
        isi=13
    else if (isi="prec in sight 14")
        isi=14
    else if (isi="prec in sight 15")
        isi=15
    else if (isi="prec in sight 16")
        isi=16
    else if (isi="ts no prec")
        isi=17
    else if (isi="squalls")
        isi=18
    else if (isi="funnel cld")
        isi=19
    else if (isi="re dz")
        isi=20
    else if (isi="re ra (not fr)")
        isi=21
    else if (isi="re ra")
        isi=21
    else if (isi="re fr dz")
        isi=24
    else if (isi="re fr ra")
        isi=24
    else if (isi="re sh of ra")
        isi=25
    else if (isi="re fog")
        isi=28
    else if (isi="re ts")
        isi=29
    else if (isi="inter sl ra")
        isi=60
    else if (isi="cns sl ra")
        isi=61
    else if (isi="inter mod ra")
        isi=62
    else if (isi="cns mod ra")
        isi=63
    else if (isi="inter heavy ra")
        isi=64
    else if (isi="cns heavy ra")
        isi=65
    else if (isi="sl ra fr")
        isi=66
    else if (isi="mod/heavy ra fr")
        isi=67
    else if (isi="sl ra sh")
        isi=80
    else if (isi="sl ra re ts")
        isi=91
    else if (isi="mod/heavy ra re ts")
        isi=92
    else if (isi="mod ra re ts")
        isi=92
    else if (isi="heavy ra re ts")
        isi=92
    else if (isi="sl ts no ha+ra")
        isi=95
    else if (isi="mod ts no ha+ra")
        isi=95
    else if (isi="sl/mod ts no ha+ra")
        isi=95
    else if (isi="sl ts+ra")
        isi=95
    else if (isi="sl ts + ra")
        isi=95
    else if (isi="mod ts+ra")
        isi=95
    else if (isi="mod ts + ra")
        isi=95
    else if (isi="sl/mod ts + ra")
        isi=95
    else if (isi="heavy ts no ha+ra")
        isi=97
    else if (isi="heavy ts+ra")
        isi=97
    else if (isi="heavy ts + ra")
        isi=97
    
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

w1:=12+space
isi:= wbc.Range("I"w1).Value
if(isi != ""){
    if (isi="Cloudy -")
        isi=0
    else if (isi="Cloudy +-")
        isi=1
    else if (isi="Cloudy+")
        isi=2
    else if (isi="Cloudy +")
        isi=2
    else if (isi="Sand")
        isi=3
    else if (isi="haze")
        isi=4
    else if (isi="Dz")
        isi=5
    else if (isi="Drizzel")
        isi=5
    else if (isi="Sh")
        isi=8
    else if (isi="ra")
        isi=6
    else if (isi="snow")
        isi=7
    else if (isi="Shower")
        isi=8
    else if (isi="Ts")
        isi=9
    else if (isi="Thunderstorm")
        isi=9

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

w2:=13+space
isi:= wbc.Range("I"w2).Value
if(isi != ""){
    if (isi="Cloudy -")
        isi=0
    else if (isi="Cloudy +-")
        isi=1
    else if (isi="Cloudy+")
        isi=2
    else if (isi="Cloudy +")
        isi=2
    else if (isi="Sand")
        isi=3
    else if (isi="haze")
        isi=4
    else if (isi="Dz")
        isi=5
    else if (isi="Drizzel")
        isi=5
    else if (isi="Sh")
        isi=8
    else if (isi="ra")
        isi=6
    else if (isi="snow")
        isi=7
    else if (isi="Shower")
        isi=8
    else if (isi="Ts")
        isi=9
    else if (isi="Thunderstorm")
        isi=9
        
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

degPanas:=10+space
isi:= wbc.Range("J"degPanas).Value
SendInput, % isi
Send {Tab}

pDibaca:=11+space
isi:= wbc.Range("J"pDibaca).Value
SendInput, % isi
Send {Tab}

if(Mod(jam,3)=0){
    deltaP:=13+space
    isi:= wbc.Range("K"deltaP).Value
    SendInput, % isi
    Send {Tab}
}

qff:=11+space
isi:= wbc.Range("K"qff).Value
SendInput, % isi
Send {Tab}

qfe:=11+space
isi:= wbc.Range("L"qfe).Value
SendInput, % isi
Send {Tab}

if(Mod(jam,12)=0){
    deltaP24:=13+space
    isi:= wbc.Range("L"deltaP24).Value
    SendInput, % isi
    Send {Tab}
}

bolaKering:=10+space
isi:= wbc.Range("M"bolaKering).Value
SendInput, % isi
Send {Tab}

bolaBasah:=12+space
isi:= wbc.Range("M"bolaBasah).Value
SendInput, % isi
Send {Tab}

dp:=10+space
isi:= wbc.Range("N"dp).Value
SendInput, % isi
Send {Tab}

rh:=13+space
isi:= wbc.Range("N"rh).Value
SendInput, % isi
Send {Tab}

if(jam=12){
    isi:= wbc.Range("O98").Value
    SendInput, % isi
    Send {Tab}
}

if(jam=0){
    isi:= wbc.Range("O12").Value
    SendInput, % isi
    Send {Tab}
}

if(Mod(jam,3)=0){
    dataHujan:=10+space
    isiDataHujan:= wbc.Range("Q"dataHujan).Value
    if(isiDataHujan >= 0){
        Send {Down} 
        Loop, % isiDataHujan
        {
            Send {Down} 
        }
        Send {Enter}
        Send {Tab}
    }
    else
        Send {Tab}      
    if(isiDataHujan != 3){
        hujanTakaranTerakhir:= 11+space
        isi:= wbc.Range("Q"hujanTakaranTerakhir).Value
        SendInput, % isi
        Send {Tab}
    } 
}

if(Mod(jam,6)=0 and isiDataHujan != 3){
    hujan6Jam:=12+space
    isi:= wbc.Range("Q"hujan6Jam).Value
    SendInput, % isi
    Send {Tab}
}

if(jam=0){
    isi:= wbc.Range("Q13").Value
    SendInput, % isi
    Send {Tab}
}

awanRendah:=10+space
jenisAwanRendah:= wbc.Range("T"awanRendah).Value
if(jenisAwanRendah != ""){
    if (jenisAwanRendah="cu sc")
        jenisAwanRendah=8
    else if (jenisAwanRendah="cusc")
        jenisAwanRendah=8
    else if (jenisAwanRendah="cbcu")
        jenisAwanRendah=9
    else if (jenisAwanRendah="cb cu")
        jenisAwanRendah=9
    else if (jenisAwanRendah="cb cu st")
        jenisAwanRendah=9
    else if (jenisAwanRendah="acas")
        jenisAwanRendah=7
    else if (jenisAwanRendah="cb")
        jenisAwanRendah=9
    else if (jenisAwanRendah=0)
        jenisAwanRendah=10
    else if (jenisAwanRendah="ci")
        jenisAwanRendah=0
    else if (jenisAwanRendah="cc")
        jenisAwanRendah=1
    else if (jenisAwanRendah="cs")
        jenisAwanRendah=2
    else if (jenisAwanRendah="ac")
        jenisAwanRendah=3
    else if (jenisAwanRendah="as")
        jenisAwanRendah=4
    else if (jenisAwanRendah="ns")
        jenisAwanRendah=5
    else if (jenisAwanRendah="sc")
        jenisAwanRendah=6
    else if (jenisAwanRendah="st")
        jenisAwanRendah=7
    else if (jenisAwanRendah="cu")
        jenisAwanRendah=8
    else if (jenisAwanRendah="cb")
        jenisAwanRendah=9
    else{
        MsgBox, Jenis awan belum terdaftar di program `nkontak 089677030198.
        ExitApp
        return
    }    
    Send {Down}
    Loop, % jenisAwanRendah
    {
        Send {Down} 
    }
    Send {Enter}
    Send {Tab}
}
else
    Send {Tab}

if(jenisAwanRendah>0){
    tinggiDasarCL:=12+space
    isiTinggiDasarCL:= wbc.Range("T"tinggiDasarCL).Value
    split:="/"
    IfInString, isiTinggiDasarCL, %split%
    {
        for index, isiTinggiDasarCL in StrSplit(isiTinggiDasarCL, "/"){
            SendInput, % isiTinggiDasarCL
            if(index.MaxIndex()=2){
                Send {Tab 2}
            }
            else{
                Send {Tab}
            }
        }
    }
    else {
        SendInput, % isiTinggiDasarCL
        Send {Tab 3}
    }

    tinggiPuncakCL1:=13+space
    isiPuncakCL1:= wbc.Range("T"tinggiPuncakCL1).Value
    split:="/"
    IfInString, isiPuncakCL1, %split%
    {
        for index, isiPuncakCL1 in StrSplit(isiPuncakCL1, "/"){
            SendInput, % isiPuncakCL1
            Send {Tab}
        }
    }

    else {
        SendInput, % isiPuncakCL1
        Send {Tab 2}
    }

    arahElevasiCL1:=10+space
    isi:= wbc.Range("W"arahElevasiCL1).Value
    if(isi != ""){
        if (isi="stnr")
            isi=0
        else if (isi="north east")
            isi=1
        else if (isi="east")
            isi=2
        else if (isi="south east")
            isi=3
        else if (isi="south")
            isi=4
        else if (isi="south west")
            isi=5
        else if (isi="west")
            isi=6
        else if (isi="north west")
            isi=7
        else if (isi="north")
            isi=8
        else if (isi="-")
            isi=9
        else if (isi="No cloud")
            isi=0
        else if (isi=5)
            isi=8
        else if (isi=10)
            isi=8
        else if (isi=15)
            isi=8
        else if (isi=20)
            isi=8
        else if (isi=25)
            isi=1
        else if (isi=30)
            isi=1
        else if (isi=35)
            isi=1
        else if (isi=40)
            isi=1
        else if (isi=45)
            isi=1
        else if (isi=50)
            isi=1
        else if (isi=55)
            isi=1
        else if (isi=60)
            isi=1
        else if (isi=65)
            isi=1
        else if (isi=70)
            isi=2
        else if (isi=75)
            isi=2
        else if (isi=80)
            isi=2
        else if (isi=85)
            isi=2
        else if (isi=90)
            isi=2
        else if (isi=95)
            isi=2
        else if (isi=100)
            isi=2
        else if (isi=105)
            isi=2
        else if (isi=110)
            isi=2
        else if (isi=115)
            isi=3
        else if (isi=120)
            isi=3
        else if (isi=125)
            isi=3
        else if (isi=130)
            isi=3
        else if (isi=135)
            isi=3
        else if (isi=140)
            isi=3
        else if (isi=145)
            isi=3
        else if (isi=150)
            isi=3
        else if (isi=155)
            isi=3
        else if (isi=160)
            isi=4
        else if (isi=165)
            isi=4
        else if (isi=170)
            isi=4
        else if (isi=175)
            isi=4
        else if (isi=180)
            isi=4
        else if (isi=185)
            isi=4
        else if (isi=190)
            isi=4
        else if (isi=195)
            isi=4
        else if (isi=200)
            isi=4
        else if (isi=205)
            isi=5
        else if (isi=210)
            isi=5
        else if (isi=215)
            isi=5
        else if (isi=220)
            isi=5
        else if (isi=225)
            isi=5
        else if (isi=230)
            isi=5
        else if (isi=235)
            isi=5
        else if (isi=240)
            isi=5
        else if (isi=245)
            isi=5
        else if (isi=250)
            isi=6
        else if (isi=255)
            isi=6
        else if (isi=260)
            isi=6
        else if (isi=265)
            isi=6
        else if (isi=270)
            isi=6
        else if (isi=275)
            isi=6
        else if (isi=280)
            isi=6
        else if (isi=285)
            isi=6
        else if (isi=290)
            isi=6
        else if (isi=295)
            isi=7
        else if (isi=300)
            isi=7
        else if (isi=305)
            isi=7
        else if (isi=310)
            isi=7
        else if (isi=315)
            isi=7
        else if (isi=320)
            isi=7
        else if (isi=325)
            isi=7
        else if (isi=330)
            isi=7
        else if (isi=335)
            isi=7
        else if (isi=340)
            isi=8
        else if (isi=345)
            isi=8
        else if (isi=350)
            isi=8
        else if (isi=355)
            isi=8
        else if (isi=360)
            isi=8
        Send {Down 2}
        Loop, % isi
        {
            Send {Down} 
        }
        Send {Enter}
        Send {Tab}
    }
    else
        Send {Tab}

    isi:= wbc.Range("W"arahElevasiCL1).Value
    if(isi != ""){
        if (isi="stnr")
            isi=0
        else if (isi="north east")
            isi=1
        else if (isi="east")
            isi=2
        else if (isi="south east")
            isi=3
        else if (isi="south")
            isi=4
        else if (isi="south west")
            isi=5
        else if (isi="west")
            isi=6
        else if (isi="north west")
            isi=7
        else if (isi="north")
            isi=8
        else if (isi="-")
            isi=9
        else if (isi="No cloud")
            isi=0
        else if (isi=5)
            isi=8
        else if (isi=10)
            isi=8
        else if (isi=15)
            isi=8
        else if (isi=20)
            isi=8
        else if (isi=25)
            isi=1
        else if (isi=30)
            isi=1
        else if (isi=35)
            isi=1
        else if (isi=40)
            isi=1
        else if (isi=45)
            isi=1
        else if (isi=50)
            isi=1
        else if (isi=55)
            isi=1
        else if (isi=60)
            isi=1
        else if (isi=65)
            isi=1
        else if (isi=70)
            isi=2
        else if (isi=75)
            isi=2
        else if (isi=80)
            isi=2
        else if (isi=85)
            isi=2
        else if (isi=90)
            isi=2
        else if (isi=95)
            isi=2
        else if (isi=100)
            isi=2
        else if (isi=105)
            isi=2
        else if (isi=110)
            isi=2
        else if (isi=115)
            isi=3
        else if (isi=120)
            isi=3
        else if (isi=125)
            isi=3
        else if (isi=130)
            isi=3
        else if (isi=135)
            isi=3
        else if (isi=140)
            isi=3
        else if (isi=145)
            isi=3
        else if (isi=150)
            isi=3
        else if (isi=155)
            isi=3
        else if (isi=160)
            isi=4
        else if (isi=165)
            isi=4
        else if (isi=170)
            isi=4
        else if (isi=175)
            isi=4
        else if (isi=180)
            isi=4
        else if (isi=185)
            isi=4
        else if (isi=190)
            isi=4
        else if (isi=195)
            isi=4
        else if (isi=200)
            isi=4
        else if (isi=205)
            isi=5
        else if (isi=210)
            isi=5
        else if (isi=215)
            isi=5
        else if (isi=220)
            isi=5
        else if (isi=225)
            isi=5
        else if (isi=230)
            isi=5
        else if (isi=235)
            isi=5
        else if (isi=240)
            isi=5
        else if (isi=245)
            isi=5
        else if (isi=250)
            isi=6
        else if (isi=255)
            isi=6
        else if (isi=260)
            isi=6
        else if (isi=265)
            isi=6
        else if (isi=270)
            isi=6
        else if (isi=275)
            isi=6
        else if (isi=280)
            isi=6
        else if (isi=285)
            isi=6
        else if (isi=290)
            isi=6
        else if (isi=295)
            isi=7
        else if (isi=300)
            isi=7
        else if (isi=305)
            isi=7
        else if (isi=310)
            isi=7
        else if (isi=315)
            isi=7
        else if (isi=320)
            isi=7
        else if (isi=325)
            isi=7
        else if (isi=330)
            isi=7
        else if (isi=335)
            isi=7
        else if (isi=340)
            isi=8
        else if (isi=345)
            isi=8
        else if (isi=350)
            isi=8
        else if (isi=355)
            isi=8
        else if (isi=360)
            isi=8
        Send {Down 2}
        Loop, % isi
        {
            Send {Down} 
        }
        Send {Enter}
        Send {Tab}
    }
    else
        Send {Tab}

    nAwan:=12+space
    isi:= wbc.Range("W"nAwan).Value
    if(isi>=0){
        Send {Down}
        Loop, % isi
        {
            Send {Down} 
        }
        Send {Enter}
        Send {Tab}
    }
    else{
        Send {Tab}
    }

    sudutCL:=11+space
    pilihSudut:= wbc.Range("W"sudutCL).Value
if(pilihSudut !=""){
    split:="/"
    IfInString, pilihSudut, %split%
    {
        for index, pilihSudut in StrSplit(pilihSudut, "/"){
            if(pilihSudut>=45){
                pilihSudut=1
            }
            else if(pilihSudut=30){
                pilihSudut=2
            }
            else if(pilihSudut=20){
                pilihSudut=3
            }
            else if(pilihSudut=15){
                pilihSudut=4
            }
            else if(pilih=12){
                pilihSudut=5
            }
            else if(pilihSudut=9){
                pilihSudut=6
            }
            else if(pilihSudut=7){
                pilihSudut=7
            }
            else if(pilihSudut=6){
                pilihSudut=8
            }
            else if(pilihSudut<6 and pilihSudut>0){
                pilihSudut=9
            }
            else{
                MsgBox, Sudut belum terdaftar di program `nkontak 089677030198.     
                ExitApp
                return
            }  
            if(pilihSudut>=0){
                Send {Down 2}
                Loop, % pilihSudut
                {
                    Send {Down} 
                }
                Send {Enter}
                Send {Tab}
            }
        }
    }
    else{
        if(pilihSudut>=45){
                pilihSudut=1
            }
        else if(pilihSudut=30){
            pilihSudut=2
        }
        else if(pilihSudut=20){
            pilihSudut=3
        }
        else if(pilihSudut=15){
            pilihSudut=4
        }
        else if(pilih=12){
            pilihSudut=5
        }
        else if(pilihSudut=9){
            pilihSudut=6
        }
        else if(pilihSudut=7){
            pilihSudut=7
        }
        else if(pilihSudut=6){
            pilihSudut=8
        }
        else if(pilihSudut<6){
            pilihSudut=9
        }
        Send {Down 2}
        Loop, % pilihSudut
        {
            Send {Down} 
        }
        Send {Enter}
        Send {Tab 2}
    }
    }
    else{
     Send {Tab 2}   
    }
   
}
;batas awan rendah

nAwanMenengah:=10+space
isiNAwanMenengah:= wbc.Range("Z"nAwanMenengah).Value
if(isiNAwanMenengah>=0){
    Send {Down}
    Loop, % isiNAwanMenengah
    {
        Send {Down} 
    }
    Send {Enter}
    Send {Tab}
}
else if (isiNAwanMenengah=""){
    Loop, 11
    {
        Send {Down} 
    }
    Send {Enter}
    Send {Tab}
    isiNAwanMenengah=11
}
else{
    Send {Tab}   
}

nAwanTinggi:=12+space
isiNAwanTinggi:= wbc.Range("Z"nAwanTinggi).Value
if(isiNAwanTinggi>=0){
    Send {Down}
    Loop, % isiNAwanTinggi
    {
        Send {Down} 
    }
    Send {Enter}
    Send {Tab}
}
else if (isiNAwanTinggi=""){
    Loop, 11
    {
        Send {Down} 
    }
    Send {Enter}
    Send {Tab}
    isiNAwanTinggi=11
}
else{
    Send {Tab}   
}

if(isiNAwanMenengah > 0){
    awanMenengah:=10+space
    isi:= wbc.Range("AC"awanMenengah).Value
    if(isi != ""){
        if (isi="cu sc")
            isi=8
        else if (isi="cusc")
            isi=8
        else if (isi="cbcu")
            isi=9
        else if (isi="cb cu")
            isi=9
        else if (isi="cb cu st")
            isi=9
        else if (isi="acas")
            isi=7
        else if (isi="ac as")
            isi=7
        else if (isi="cb")
            isi=9
        else if (isi="cu")
            isi=9
        else if (isi="0")
            isi=10
        else if (isi="ci")
            isi=0
        else if (isi="cc")
            isi=1
        else if (isi="cs")
            isi=2
        else if (isi="ac")
            isi=3
        else if (isi="as")
            isi=4
        else if (isi="ns")
            isi=5
        else if (isi="sc")
            isi=6
        else if (isi="st")
            isi=7
        else if (isi="cu")
            isi=8
        else if (isi="cb")
            isi=9
        else{
            MsgBox, Jenis awan belum terdaftar di program `nkontak 089677030198.     
            ExitApp
            return
        }     
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
}

if(isiNAwanTinggi>0){
    awanTinggi:=12+space
    isi:= wbc.Range("AC"awanTinggi).Value
    if(isi != ""){
        if (isi="cu sc")
            isi=8
        else if (isi="cusc")
            isi=8
        else if (isi="cbcu")
            isi=9
        else if (isi="cb cu")
            isi=9
        else if (isi="cb cu st")
            isi=9
        else if (isi="acas")
            isi=7
        else if (isi="ac as")
            isi=7
        else if (isi="cb")
            isi=9
        else if (isi="cu")
            isi=9
        else if (isi="0")
            isi=10
        else if (isi="ci")
            isi=0
        else if (isi="cc")
            isi=1
        else if (isi="cs")
            isi=2
        else if (isi="ac")
            isi=3
        else if (isi="as")
            isi=4
        else if (isi="ns")
            isi=5
        else if (isi="sc")
            isi=6
        else if (isi="st")
            isi=7
        else if (isi="cu")
            isi=8
        else if (isi="cb")
            isi=9
        else{
            MsgBox, Jenis awan belum terdaftar di program `nkontak 089677030198.     
            ExitApp
            return
        }     
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
}

if(isiNAwanMenengah>0){
    arahElevasiMenengah:=10+space
    isi:= wbc.Range("AE"arahElevasiMenengah).Value
    if(isi != ""){
        if (isi="stnr")
            isi=0
        else if (isi="north east")
            isi=1
        else if (isi="east")
            isi=2
        else if (isi="south east")
            isi=3
        else if (isi="south")
            isi=4
        else if (isi="south west")
            isi=5
        else if (isi="west")
            isi=6
        else if (isi="north west")
            isi=7
        else if (isi="north")
            isi=8
        else if (isi="-")
            isi=9
        else if (isi="No cloud")
            isi=0
        else if (isi=5)
            isi=8
        else if (isi=10)
            isi=8
        else if (isi=15)
            isi=8
        else if (isi=20)
            isi=8
        else if (isi=25)
            isi=1
        else if (isi=30)
            isi=1
        else if (isi=35)
            isi=1
        else if (isi=40)
            isi=1
        else if (isi=45)
            isi=1
        else if (isi=50)
            isi=1
        else if (isi=55)
            isi=1
        else if (isi=60)
            isi=1
        else if (isi=65)
            isi=1
        else if (isi=70)
            isi=2
        else if (isi=75)
            isi=2
        else if (isi=80)
            isi=2
        else if (isi=85)
            isi=2
        else if (isi=90)
            isi=2
        else if (isi=95)
            isi=2
        else if (isi=100)
            isi=2
        else if (isi=105)
            isi=2
        else if (isi=110)
            isi=2
        else if (isi=115)
            isi=3
        else if (isi=120)
            isi=3
        else if (isi=125)
            isi=3
        else if (isi=130)
            isi=3
        else if (isi=135)
            isi=3
        else if (isi=140)
            isi=3
        else if (isi=145)
            isi=3
        else if (isi=150)
            isi=3
        else if (isi=155)
            isi=3
        else if (isi=160)
            isi=4
        else if (isi=165)
            isi=4
        else if (isi=170)
            isi=4
        else if (isi=175)
            isi=4
        else if (isi=180)
            isi=4
        else if (isi=185)
            isi=4
        else if (isi=190)
            isi=4
        else if (isi=195)
            isi=4
        else if (isi=200)
            isi=4
        else if (isi=205)
            isi=5
        else if (isi=210)
            isi=5
        else if (isi=215)
            isi=5
        else if (isi=220)
            isi=5
        else if (isi=225)
            isi=5
        else if (isi=230)
            isi=5
        else if (isi=235)
            isi=5
        else if (isi=240)
            isi=5
        else if (isi=245)
            isi=5
        else if (isi=250)
            isi=6
        else if (isi=255)
            isi=6
        else if (isi=260)
            isi=6
        else if (isi=265)
            isi=6
        else if (isi=270)
            isi=6
        else if (isi=275)
            isi=6
        else if (isi=280)
            isi=6
        else if (isi=285)
            isi=6
        else if (isi=290)
            isi=6
        else if (isi=295)
            isi=7
        else if (isi=300)
            isi=7
        else if (isi=305)
            isi=7
        else if (isi=310)
            isi=7
        else if (isi=315)
            isi=7
        else if (isi=320)
            isi=7
        else if (isi=325)
            isi=7
        else if (isi=330)
            isi=7
        else if (isi=335)
            isi=7
        else if (isi=340)
            isi=8
        else if (isi=345)
            isi=8
        else if (isi=350)
            isi=8
        else if (isi=355)
            isi=8
        else if (isi=360)
            isi=8
        Send {Down 2}
        Loop, % isi
        {
            Send {Down} 
        }
        Send {Enter}
        Send {Tab}
    }
    else
        Send {Tab}
            
    tinggiAwanMenegah:=11+space
    isi:= wbc.Range("AE"tinggiAwanMenegah).Value
    SendInput, % Floor(isi)
    Send {Tab}
}

if(isiNAwanTinggi>0){
    arahElevasiTinggi:=12+space
    isi:= wbc.Range("AE"arahElevasiMenengah).Value
    if(isi != ""){
        if (isi="stnr")
            isi=0
        else if (isi="north east")
            isi=1
        else if (isi="east")
            isi=2
        else if (isi="south east")
            isi=3
        else if (isi="south")
            isi=4
        else if (isi="south west")
            isi=5
        else if (isi="west")
            isi=6
        else if (isi="north west")
            isi=7
        else if (isi="north")
            isi=8
        else if (isi="-")
            isi=9
        else if (isi="No cloud")
            isi=0
        else if (isi=5)
            isi=8
        else if (isi=10)
            isi=8
        else if (isi=15)
            isi=8
        else if (isi=20)
            isi=8
        else if (isi=25)
            isi=1
        else if (isi=30)
            isi=1
        else if (isi=35)
            isi=1
        else if (isi=40)
            isi=1
        else if (isi=45)
            isi=1
        else if (isi=50)
            isi=1
        else if (isi=55)
            isi=1
        else if (isi=60)
            isi=1
        else if (isi=65)
            isi=1
        else if (isi=70)
            isi=2
        else if (isi=75)
            isi=2
        else if (isi=80)
            isi=2
        else if (isi=85)
            isi=2
        else if (isi=90)
            isi=2
        else if (isi=95)
            isi=2
        else if (isi=100)
            isi=2
        else if (isi=105)
            isi=2
        else if (isi=110)
            isi=2
        else if (isi=115)
            isi=3
        else if (isi=120)
            isi=3
        else if (isi=125)
            isi=3
        else if (isi=130)
            isi=3
        else if (isi=135)
            isi=3
        else if (isi=140)
            isi=3
        else if (isi=145)
            isi=3
        else if (isi=150)
            isi=3
        else if (isi=155)
            isi=3
        else if (isi=160)
            isi=4
        else if (isi=165)
            isi=4
        else if (isi=170)
            isi=4
        else if (isi=175)
            isi=4
        else if (isi=180)
            isi=4
        else if (isi=185)
            isi=4
        else if (isi=190)
            isi=4
        else if (isi=195)
            isi=4
        else if (isi=200)
            isi=4
        else if (isi=205)
            isi=5
        else if (isi=210)
            isi=5
        else if (isi=215)
            isi=5
        else if (isi=220)
            isi=5
        else if (isi=225)
            isi=5
        else if (isi=230)
            isi=5
        else if (isi=235)
            isi=5
        else if (isi=240)
            isi=5
        else if (isi=245)
            isi=5
        else if (isi=250)
            isi=6
        else if (isi=255)
            isi=6
        else if (isi=260)
            isi=6
        else if (isi=265)
            isi=6
        else if (isi=270)
            isi=6
        else if (isi=275)
            isi=6
        else if (isi=280)
            isi=6
        else if (isi=285)
            isi=6
        else if (isi=290)
            isi=6
        else if (isi=295)
            isi=7
        else if (isi=300)
            isi=7
        else if (isi=305)
            isi=7
        else if (isi=310)
            isi=7
        else if (isi=315)
            isi=7
        else if (isi=320)
            isi=7
        else if (isi=325)
            isi=7
        else if (isi=330)
            isi=7
        else if (isi=335)
            isi=7
        else if (isi=340)
            isi=8
        else if (isi=345)
            isi=8
        else if (isi=350)
            isi=8
        else if (isi=355)
            isi=8
        else if (isi=360)
            isi=8
        Send {Down 2}
        Loop, % isi
        {
            Send {Down} 
        }
        Send {Enter}
        Send {Tab}
    }
    else
        Send {Tab}

    tinggiAwanTinggi:=13+space
    isi:= wbc.Range("AE"tinggiAwanTinggi).Value
    SendInput, % Floor(isi)
    Send {Tab}
}

nLangitTertutup:=10+space
isi:= wbc.Range("AG"nLangitTertutup).Value
if(isi>=0){
    Send {Down}
    Loop, % isi
    {
        Send {Down} 
    }
    Send {Enter}
    Send {Tab}
}
else if (isi=""){
    Loop, 11
    {
        Send {Down} 
    }
    Send {Enter}
    Send {Tab}
}
else
    Send {Tab}

jenisC1:=10+space
isi:= wbc.Range("AI"jenisC1).Value
if(isi != "-" and isi != ""){
    if (isi="cu sc")
        isi=8
    else if (isi="cusc")
        isi=8
    else if (isi="cbcu")
        isi=9
    else if (isi="cb cu")
        isi=9
    else if (isi="cb cu st")
        isi=9
    else if (isi="acas")
        isi=7
else if (isi="ac as")
        isi=7
    else if (isi="cb")
        isi=9
    else if (isi="0")
        isi=10
    else if (isi="ci")
        isi=0
    else if (isi="cc")
        isi=1
    else if (isi="cs")
        isi=2
    else if (isi="ac")
        isi=3
    else if (isi="as")
        isi=4
    else if (isi="ns")
        isi=5
    else if (isi="sc")
        isi=6
    else if (isi="st")
        isi=7
    else if (isi="cu")
        isi=8
    else if (isi="cb")
        isi=9 
    else{
            MsgBox, Jenis awan belum terdaftar di program `nkontak 089677030198.     
            ExitApp
            return
        } 
    Send {Down 2}
    Loop, % isi
    {
        Send {Down} 
    }
    Send {Enter}
    Send {Tab}
}
else if (isi="-"){
    Send {Down}
    Send {Enter}
    Send {Tab}
}
else
    Send {Tab}

jenisC2:=11+space
isi:= wbc.Range("AI"jenisC2).Value
if(isi != "-" and isi != ""){ 
    if (isi="cu sc")
        isi=8
    else if (isi="cusc")
        isi=8
    else if (isi="cbcu")
        isi=9
    else if (isi="cb cu")
        isi=9
    else if (isi="cb cu st")
        isi=9
    else if (isi="acas")
        isi=7
else if (isi="ac as")
        isi=7
    else if (isi="cb")
        isi=9
    else if (isi="0")
        isi=10
    else if (isi="ci")
        isi=0
    else if (isi="cc")
        isi=1
    else if (isi="cs")
        isi=2
    else if (isi="ac")
        isi=3
    else if (isi="as")
        isi=4
    else if (isi="ns")
        isi=5
    else if (isi="sc")
        isi=6
    else if (isi="st")
        isi=7
    else if (isi="cu")
        isi=8
    else if (isi="cb")
        isi=9 
    else{
            MsgBox, Jenis awan belum terdaftar di program `nkontak 089677030198.     
            ExitApp
            return
        } 
    Send {Down 2}
    Loop, % isi
    {
        Send {Down} 
    }
    Send {Enter}
    Send {Tab}
}
else if (isi="-"){
    Send {Down}
    Send {Enter}
    Send {Tab}
}
else
    Send {Tab}

jenisC3:=12+space
isi:= wbc.Range("AI"jenisC3).Value
if(isi != "-" and isi != ""){
    if (isi="cu sc")
        isi=8
    else if (isi="cusc")
        isi=8
    else if (isi="cbcu")
        isi=9
    else if (isi="cb cu")
        isi=9
    else if (isi="cb cu st")
        isi=9
    else if (isi="acas")
        isi=7
else if (isi="ac as")
        isi=7
    else if (isi="cb")
        isi=9
    else if (isi="0")
        isi=10
    else if (isi="ci")
        isi=0
    else if (isi="cc")
        isi=1
    else if (isi="cs")
        isi=2
    else if (isi="ac")
        isi=3
    else if (isi="as")
        isi=4
    else if (isi="ns")
        isi=5
    else if (isi="sc")
        isi=6
    else if (isi="st")
        isi=7
    else if (isi="cu")
        isi=8
    else if (isi="cb")
        isi=9 
    else{
            MsgBox, Jenis awan belum terdaftar di program `nkontak 089677030198.     
            ExitApp
            return
        } 
    Send {Down 2}
    Loop, % isi
    {
        Send {Down} 
    }
    Send {Enter}
    Send {Tab 2}
}
else if (isi="-"){
    Send {Down}
    Send {Enter}
    Send {Tab 2}
}
else
    Send {Tab 2}

tinggiLapisan1:=10+space
isi:= wbc.Range("AL"tinggiLapisan1).Value
SendInput, % Floor(isi)
Send {Tab}

tinggiLapisan2:=11+space
isi:= wbc.Range("AL"tinggiLapisan2).Value
SendInput, % Floor(isi)
Send {Tab}

tinggiLapisan3:=12+space
isi:= wbc.Range("AL"tinggiLapisan3).Value
SendInput, % Floor(isi)
Send {Tab 2}

banyakLapisan1:=10+space
isi:= wbc.Range("AO"banyakLapisan1).Value
if(isi>=0){
    Send {Down}
    Loop, % isi
    {
        Send {Down} 
    }
    Send {Enter}
    Send {Tab}
}
else if (isi="")
    Send {Tab}
else 
{
    Loop, 11
    {
        Send {Down} 
    }
    Send {Enter}
    Send {Tab}
}

banyakLapisan2:=11+space
isi:= wbc.Range("AO"banyakLapisan2).Value
if(isi>=0){
    Send {Down}
    Loop, % isi
    {
        Send {Down} 
    }
    Send {Enter}
    Send {Tab}
}
else if (isi="")
    Send {Tab}
else{
    Loop, 11
    {
        Send {Down} 
    }
    Send {Enter}
    Send {Tab}
}


banyakLapisan3:=12+space
isi:= wbc.Range("AO"banyakLapisan3).Value
if(isi>=0){
    Send {Down}
    Loop, % isi
    {
        Send {Down} 
    }
    Send {Enter}
    Send {Tab 2}
}
else if (isi="")
    Send {Tab 2}
else{
    Loop, 11
    {
        Send {Down} 
    }
    Send {Enter}
    Send {Tab 2}
}

if(jam=0){
    isi:= wbc.Range("AQ10").Value
    if(isi>=0){
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

    isi:= wbc.Range("AS10").Value
    SendInput, % isi
    Send {Tab} 

    isi:= wbc.Range("AQ12").Value
    SendInput, % isi
    Send {Tab}  

    isi:= wbc.Range("AQ13").Value
    SendInput, % isi
    Send {Tab}     
}

keadaanTanah:=10+space
isi:= wbc.Range("AV"keadaanTanah).Value
if(isi>=0){
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

MsgBox Periksa lalu tekan commit ^^ `n ~ PKL PENS 2021

return
