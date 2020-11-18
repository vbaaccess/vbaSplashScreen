Option Compare Database
Option Explicit

Private Const CurrentModuleName = "clsSplashScreen"

Private Sub Class_Initialize()
    '--- rozpoczecie uruchamiania ---
    Debug.Print "Konstruktor klasy SplashScreen do SURa (" & VBA.Now & ") - rozpoczecie procesu inicjowania"
End Sub

Private Sub Class_Terminate()
    '--- koniec uruchamiania ---
    Debug.Print "Destruktor klasy SplashScreen do SURa (" & VBA.Now & ") - zakończenie procesu inicjowania"
End Sub

Public Sub INIT()
    Debug.Print "Wykonanie czynnosci pomiedzy logowaniem a formularzem glownym (fLogowanie=>fGlowny)"
End Sub
