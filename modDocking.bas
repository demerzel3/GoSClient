Attribute VB_Name = "modDocking"
Option Explicit

'Impostando uno (o piu') di questi flag al momento dell'Init della propria cFine,
'una form puo' decidere quale input ricevera' dalla classe stessa sotto forma di
'eventi; secondo quanto segue:

'la finestra non riceve nessun evento supplementare
'questo flag non deve essere associato a nessun altro
Global Const WINREC_NONE As Integer = &H0

'la finestra riceve l'output del mud
Global Const WINREC_OUTPUT As Integer = &H1

'la finestra riceve (e puo' bloccare) l'input dell'utente
Global Const WINREC_INPUT As Integer = &H10

'la finestra riceve tutti i dati inviati dalla base
Global Const WINREC_ALL As Integer = WINREC_OUTPUT Or WINREC_INPUT

'tipo di output
Global Const TOUT_SOCKET As Integer = 0
Global Const TOUT_ERROR As Integer = 1
Global Const TOUT_SERVICE As Integer = 2
Global Const TOUT_STATUS As Integer = TOUT_SERVICE
Global Const TOUT_LASTLINE As Integer = 3
Global Const TOUT_CLEAN As Integer = 4

'tipo di input
Global Const TIN_TEXTBOX As Integer = 0
Global Const TIN_SENT As Integer = 1
Global Const TIN_TOSEND As Integer = 2
Global Const TIN_BUTTONS As Integer = 3
Global Const TIN_TOQUEUE As Integer = 4
Global Const TIN_TOQUEUECR As Integer = 5
Global Const TIN_TOQUEUEALIAS As Integer = 6
