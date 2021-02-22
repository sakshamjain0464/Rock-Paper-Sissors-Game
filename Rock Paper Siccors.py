import random                      #Imported Random for Computer Choosing
import  win32com.client                  #Imported win32com.client for Speak Function
import time as t                   #Imported timeTo define Break
import datetime                    #For Datein Data

#Defining Speak Function
s= win32com.client.Dispatch("SAPI.Spvoice")
def data(y):
    u=datetime.date.today()
    f = open("Play.Data.txt", "a")
    f.write(f"On {u} You {data}")
    f.close
    pass

L = ["Paper", "Rock", "Sissors"]
if __name__ == "__main__":
    print("Do You Want To Play With Me\nPress Enter")
    s.speak("Do You Want To Play With Me\nPress Enter")
    input()
    while(True):
        q=print("What Will You Choose\nPress r For Rock\nPress p for Paper\nPress s For Sissors")
        s.speak("What Will You Choose\nPress r For Rock\nPress p for Paper\nPress s For Sissors")
        A = input()
        if A=="r":
            Choice = random.choice(L)
            print(Choice)
            s.speak(Choice)
            t.sleep(1.0)
            if Choice=="Paper":
                print("Sorry!, You Lost")
                s.speak("Sorry!, You Lost")
                data("Lost")
                t.sleep(1.0)
            elif Choice=="Sissors":
                print("Congratulations!!!!!!!!!!!! You Won")
                s.speak("Congratulations!!!!!!!!!!!! You Won")
                data("Won")
                t.sleep(1.0)
            elif Choice=="Rock":
                print("Oh Tie")
                s.speak("Oh! Tie")
                data("Tied")
                t.sleep(1.0)
        elif A=="p":
            Choice = random.choice(L)
            print(Choice)
            s.speak(Choice)
            t.sleep(1.0)
            if Choice=="Sissors":
                print("Sorry!, You Lost")
                s.speak("Sorry!, You Lost")
                data("Lost")
                t.sleep(1.0)
            elif Choice=="Rock":
                print("Congratulations!!!!!!!!!!!! You Won")
                s.speak("Congratulations!!!!!!!!!!!! You Won")
                data("Lost")
                t.sleep(1.0)
            elif Choice=="Paper":
                print("Oh Tie")
                s.speak("Oh! Tie")
                data("Tied")
                t.sleep(1.0)
        elif A=="s":
            Choice = random.choice(L)
            print(Choice)
            s.speak(Choice)
            t.sleep(1.0)
            if Choice == "Rock":
                print("Sorry!, You Lost")
                s.speak("Sorry!, You Lost")
                data("Lost")
                t.sleep(1.0)
            elif Choice == "Paper":
                print("Congratulations!!!!!!!!!!!! You Won")
                s.speak("Congratulations!!!!!!!!!!!! You Won")
                data("Won")
                t.sleep(1.0)
            elif Choice == "Sissors":
                print("Oh Tie")
                s.speak("Oh! Tie")
                data("Tied")
                t.sleep(1.0)
        else:
            print("Sorry Wrong Choice")
            s.speak("Sorry!! Wrong Choice")
            t.sleep(1.0)
        print("Do You Want To Play More\nPress any Key To Continue\nPress 2 To Quit")
        s.speak("Do You Want To Play More\nPress any Key To Continue\nPress 2 To Quit")
        h=input()
        if h=="2":
            print("Thanks For playing with me")
            s.speak("Thanks For playing with me!!")
            break
        else:
            continue
print("Projct No. 5 Made By Saksham Jain")
s.speak("I am project no.5 made by saksham jain")
