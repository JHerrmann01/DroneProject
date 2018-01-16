#include <Servo.h> //accesses the Arduino Servo Library

Servo myservo;  // creates servo object to control a servo

int val;    // variable to read the value from the analog pin
int c = 0;
void setup()
{
  Serial.begin(9600);
  myservo.attach(10);  // ensures output to servo on pin 9
}


void loop() 
{ 
  String a = Serial.readString();
  if(a != "")
    c = a.toInt();
  Serial.println(c);
  //val = analogRead(1);  // reads the value of the potentiometer from A1 (value between 0 and 1023) 
  //Serial.println(val);
  //val = map(val, 0, 1023, 0, 180);     // converts reading from potentiometer to an output value in degrees of rotation that the servo can understand 
  myservo.write(c);                  // sets the servo position according to the input from the potentiometer 
  //Serial.println(val);
  delay(1000);     // waits 15ms for the servo to get to set position  
  
} 
