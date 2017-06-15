/*
 *Chase Byers
 * project 6
*/
#include "simpletools.h"                      // Include simple tools
const int TRIGGER_PIN = 0;
const int ECHO_PIN = 1;

int main()                                    // Main function
{
 long duration;
 long distance;
 
 low(TRIGGER_PIN);
 low(ECHO_PIN);
 pause(250);
 
  while(1)
  {
     pulse_out();
    duration
    print("");
    distance
    print("Distance =");
    pause(250);
  }  
}
