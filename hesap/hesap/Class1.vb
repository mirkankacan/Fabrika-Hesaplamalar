Public Class Class1
    Dim maaliyet1, maaliyet2, maaliyet3, maaliyet4, toplam As New Double
    Function hesapla(ByVal mal As Double, ByVal ham As Double, ByVal fin As Double, ByVal fire As Double, ByVal oran As Double, ByVal maka As Double, ByVal makç As Double, ByVal makh As Double, ByVal ürün As Double, ByVal brüt As Double, ByVal ssk As Double, ByVal yem As Double, ByVal iş As Double, ByVal maaş As Double, ByVal adet As Double)
        maaliyet1 = mal + ham + fin + fire + oran
        maaliyet2 = maka + makç + makh + ürün
        maaliyet3 = brüt + ssk + yem + iş
        maaliyet4 = maaş + adet
        toplam = maaliyet1 + maaliyet2 + maaliyet3 + maaliyet4
        Return toplam
    End Function

End Class
