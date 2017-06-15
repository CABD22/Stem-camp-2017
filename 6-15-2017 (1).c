#include "simpletools.h"                      

int main()                                   
{    
  
  int timeCounter;
  char playAgain = 'Y';
  int interval;
  float averageScore = 0;
  int rounds = 0;
  int best = 1500;
  
  srand(CNT);
    
   print("Game Instructions: \n");
   print("The red LED will turn on when the game starts. \n");
   print("When the yellow LED turns on, \n");
   print("Let go of the button as fast as you can. \n \n");
   
   while(playAgain == 'Y' || playAgain == 'y')
   {
     print("Press and hold the button to start. \n");
     while( input(0) == 0)
     {
       //intentionlly empty loop
     }
    
    high(15);
    low(14);
    
    interval = (rand() % 1001) + 500;
    pause(interval);
    
    low(15);
    high(14);
    
    timeCounter = 0;
    while( input(0) == 1)
    {
      pause(1);
      timeCounter = timeCounter + 1;
    }
    
    if (timeCounter > 0)
   {
    print("Your time was %d ms. \n", timeCounter);
    averageScore = averageScore + timeCounter;
    rounds = rounds + 1;
    
    if(timeCounter < timeCounter){
      best =  timeCounter;
    }
   }                         
   else
   {
     print("Oops! You let go of the button too soon.\n");
   }
  print("Would you like to play again? (Y/N) ");
  scan("%c", &playAgain);
  }      //end of while loop for play again?          


  averageScore = averageScore / rounds ;
  print("Goodbye! Your average score was %0.2f ms. \n", averageScore);
  print("The best score is %d ms. \n", best);
  
  return 0;
}  
  
   
      
