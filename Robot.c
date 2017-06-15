/*
 *Chase Byers
 *Control a servo motor with 2 buttons
 *program 7
 *6/15/17
*/
#include "simpletools.h"                      // Include simple tools
#include "servo.h"

const int RIGHT = 4;
const int LEFT = 3;
const int SERVO = 14;

const int RSTOP = 1700;
const int LSTOP = 10;


int main()                                    // Main function
{
  // Add startup code here.
  int angle = 900;

 
  while(1)
  {
    // Add main loop code here.
    if (input(LEFT) == 1)
    {
      angle = angle - 18;
    } 
    if (input(RIGHT) == 1)
         
    {
      angle = angle + 18;
    }
    if (angle > RSTOP);
    {
      angle = RSTOP;
    }
    if (angle < LSTOP)      
    {
      angle = LSTOP;
    }
    print("%c angle = %d %c", HOME, angle, CLREOL);
    
    servo_angle(SERVO, angle);
    pause(25);      
  } 
}