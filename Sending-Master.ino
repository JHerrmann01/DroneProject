//Sending - Master Code for Drone Project (Winter Break 2017)//

//Joystick Software//
//Right Joystick
const int Right_X_pin = 0;
const int Right_Y_pin = 1;

//Left Joystick
const int Left_X_pin = 2;
const int Left_Y_pin = 3;

int motorPower1 = 0;
int motorPower2 = 0;
int motorPower3 = 0;
int motorPower4 = 0;
int const incrementPowerValue = 5;

//Communications Software//
// Include the RFM69 and SPI libraries:
#include <RFM69.h>
#include <SPI.h>
// Addresses for this node. CHANGE THESE FOR EACH NODE!
#define NETWORKID     100   // Must be the same for all nodes (0 to 255)
#define MYNODEID      1   // My node ID (0 to 255)
#define TONODEID      2   // Destination node ID (0 to 254, 255 = broadcast)
// RFM69 frequency, uncomment the frequency of your module:
#define FREQUENCY     RF69_915MHZ
// AES encryption (or not):
#define ENCRYPT       false // Set to "true" to use encryption
#define ENCRYPTKEY    "TOPSECRETPASSWRD" // Use the same 16-byte key on all nodes
// Use ACKnowledge when sending messages (or not):
#define USEACK        true// Request ACKs or not
// Create a library object for our RFM69HCW module:
RFM69 radio;

//Final Joystick Software
String getPowerValues(){
  int RightXVal = analogRead(Right_X_pin);
  int RightYVal = analogRead(Right_Y_pin);
  int LeftXVal = analogRead(Left_X_pin);
  int LeftYVal = analogRead(Left_Y_pin);

  //Altitude Change
  if(LeftYVal >= 536)
  {
    int percentY = (.205338809)*(LeftYVal - 536)/100;
    motorPower1 += (percentY*incrementPowerValue);
    motorPower2 += (percentY*incrementPowerValue);
    motorPower3 += (percentY*incrementPowerValue);
    motorPower4 += (percentY*incrementPowerValue);
  }
  else if(LeftYVal <= 486)
  {
    int percentY = (-.2057613169)*(486-LeftYVal)/100;
    motorPower1 -= (percentY*incrementPowerValue);
    motorPower2 -= (percentY*incrementPowerValue);
    motorPower3 -= (percentY*incrementPowerValue);
    motorPower4 -= (percentY*incrementPowerValue);
  }

  if(motorPower1 > 85)
    motorPower1 = 85;
  if(motorPower2 > 85)
    motorPower2 = 85;
  if(motorPower3 > 85)
    motorPower3 = 85;
  if(motorPower4 > 85)
    motorPower4 = 85;

  //Spinning
  if(LeftXVal >= 536)
  {
    int percentX = (.205338809)*(LeftXVal - 536)/100;
    motorPower1 -= (percentX*incrementPowerValue);
    motorPower4 -= (percentX*incrementPowerValue);
  }
  else if(LeftXVal <= 486)
  {
    int percentX = (-.2057613169)*(486-LeftXVal)/100;
    motorPower2 -= (percentX*incrementPowerValue);
    motorPower3 -= (percentX*incrementPowerValue);
  }


  //X and Y Movement
  if(RightYVal >= 536)
  {
    int percentY = (.205338809)*(RightYVal - 536)/100;
    motorPower1 -= (percentY*incrementPowerValue);
    motorPower2 -= (percentY*incrementPowerValue);
  }
  else if(RightYVal <= 486)
  {
    int percentY = (-.2057613169)*(486-RightYVal)/100;
    motorPower3 -= (percentY*incrementPowerValue);
    motorPower4 -= (percentY*incrementPowerValue);
  }

  if(RightXVal >= 536)
  {
    int percentX = (.205338809)*(RightXVal - 536)/100;
    motorPower2 -= (percentX*incrementPowerValue);
    motorPower4 -= (percentX*incrementPowerValue);
  }
  else if(RightXVal <= 486)
  {
    int percentX = (-.2057613169)*(486-RightXVal)/100;
    motorPower1 -= (percentX*incrementPowerValue);
    motorPower3 -= (percentX*incrementPowerValue);
  }

  String finalString = String(motorPower1) + String(motorPower2) + String(motorPower3) + String(motorPower4);
  return finalString;
}

void setup() {
  /*
  //Joystick Initialization
  //Right JS
  pinMode(Right_X_pin, INPUT);
  pinMode(Right_Y_pin, INPUT);
  //Left JS
  pinMode(Left_X_pin, INPUT);
  pinMode(Left_Y_pin, INPUT);
  */
  //TEMP CODE FOR SERIAL CONTROL
  Serial.begin(9600);
  
  //Communications Code
  radio.initialize(FREQUENCY, MYNODEID, NETWORKID);
  radio.setHighPower(); // Always use this for RFM69HCW 

  // Turn on encryption if desired:
  if (ENCRYPT)
    radio.encrypt(ENCRYPTKEY);
}

void loop() {
  //Joystick Software
  //String finalString = getPowerValues(); //Getting the Power Values from the Joysticks
  //delay(300);
  
  //Communications Code - Sending JS Values
  static char sendbuffer[61];
  static int sendlength = 0;

  /*CODE USING JOYSTICKS - TEMPORARILY COMMENTED
  for(int x = 0; x < finalString.length();x++)
  {
    char input = finalString.charAt(x);
    if (input != '\r') // not a carriage return
    {
      sendbuffer[sendlength] = input;
      sendlength++;
    }
    //If the input is a carriage return, or the buffer is full:
    if ((input == '\r') || (sendlength == 8)) // CR or buffer full
    {
      radio.send(TONODEID, sendbuffer, sendlength);
      sendlength = 0; // reset the packet
    }
  }
  */

  if(Serial.available() > 0)
  {
    char input = Serial.read();
    
    if (input != '\r') // not a carriage return
    {
      sendbuffer[sendlength] = input;
      sendlength++;
    }

    //If the input is a carriage return, or the buffer is full:
    
    if ((input == '\r') || (sendlength == 8)) // CR or buffer full
    {
      // Send the packet!
      Serial.print("sending to node ");
      Serial.print(TONODEID, DEC);
      Serial.print(": [");
      for (byte i = 0; i < sendlength; i++)
        Serial.print(sendbuffer[i]);
      Serial.println("]");
      
      /// There are two ways to send packets. If you want
      // acknowledgements, use sendWithRetry():
      if (USEACK)
      {
        if (radio.sendWithRetry(TONODEID, sendbuffer, sendlength))
          Serial.println("ACK received!");
        else
          Serial.println("no ACK received :(");
      }

      // If you don't need acknowledgements, just use send():
      else // don't use ACK
      {
        radio.send(TONODEID, sendbuffer, sendlength);
      }
      
      sendlength = 0; // reset the packet
    }
  }

  
  
}

