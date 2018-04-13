Try 
    For i = 0
         
        dim PINGATTACCKER1 As New Threading .Thread(AdressOf PING1)
        PINGATTACCKER1.Start()
        list.Add(PINGATTACCKER1)
        If ATKSTATUS = False Then
              ATKSTATUS = False
              SubAtk = ATKVALUE
              Try
                 For x = 0 To subAtk
                     list(i).Abort()
                     list(i).Abort()
                     list(i).Abort()
                     list(i).Abort()
                     list(i).Abort()
                     list(i).Abort()
                     list(i).Abort()
                     list(i).Abort()
                     list(i).Abort()
                     list(i).Abort()
                 Next
              Catch es As Exception
              End Try
              Console.Clear()
              Console.ForegroundColor = ConsoleColor.Write
              Console.Write("Cerberus: Rapport  du Denie de service....")
              Console.Write("")
              Console.Write("Envoi de " 1 PacketValide + PacketInvalideHost + Packet***)
              Console.Write("")
              Console.Write("Requestes valides :" & PacketValide & " | Requetes invalides: *****")
              Console.Write("")
              Console.Write("Voulez-vous reinitialiser Cerberus? (O\N?")
              'Console.Read()'
Question = Console.ReadLine

Imports System.Threading
Thread.Suspend()
Thread.()