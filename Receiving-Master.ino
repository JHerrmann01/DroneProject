//Receiving - Master Code for Drone Project (Winter Break 2017)//

//Ultrasonic Sensor//
//#include <SR04.h>
//#define TRIG_PIN 13
//#define ECHO_PIN 12
//Initialization of Ultrasonic Sensor
//SR04 sr04 = SR04(ECHO_PIN,TRIG_PIN);
//long a;

//Motor Software//
#include <Servo.h> //accesses the Arduino Servo Library
Servo MotorOne;  // creates servo object to control a servo
Servo MotorTwo;  // creates servo object to control a servo
Servo MotorThree;  // creates servo object to control a servo
Servo MotorFour;  // creates servo object to control a servo
int const Motor_One_servo_pin = 8;
int const Motor_Two_servo_pin = 9;
int const Motor_Three_servo_pin = 10;
int const Motor_Four_servo_pin = 11;

//Accelerometer Software//
//#include<Wire.h>
//const int MPU_addr=0x68;  // I2C address of the MPU-6050
//int16_t AcX,AcY,AcZ,Tmp,GyX,GyY,GyZ;

//Communications Software//
// Include the RFM69 and SPI libraries:
#include <RFM69.h>
#include <SPI.h>
// Addresses for this node. CHANGE THESE FOR EACH NODE!
#define NETWORKID     100   // Must be the same for all nodes (0 to 255)
#define MYNODEID      2   // My node ID (0 to 255)
#define TONODEID      1   // Destination node ID (0 to 254, 255 = broadcast)
// RFM69 frequency, uncomment the frequency of your module:
#define FREQUENCY     RF69_915MHZ
// AES encryption (or not):
#define ENCRYPT       false // Set to "true" to use encryption
#define ENCRYPTKEY    "TOPSECRETPASSWRD" // Use the same 16-byte key on all nodes
// Use ACKnowledge when sending messages (or not):
#define USEACK        false// Request ACKs or not
// Create a library object for our RFM69HCW module:
RFM69 radio;

void setup() {
  //Motor and ESC Code
  MotorOne.attach(Motor_One_servo_pin);
  MotorTwo.attach(Motor_Two_servo_pin);
  MotorThree.attach(Motor_Three_servo_pin);
  MotorFour.attach(Motor_Four_servo_pin);  

  //Accelerometer Code
  //Wire.begin();
  //Wire.beginTransmission(MPU_addr);
  //Wire.write(0x6B);  // PWR_MGMT_1 register
  //Wire.write(0);     // set to zero (wakes up the MPU-6050)
  //Wire.endTransmission(true);
  
  //Communications Code
  radio.initialize(FREQUENCY, MYNODEID, NETWORKID);
  radio.setHighPower(); // Always use this for RFM69HCW 
  // Turn on encryption if desired:
  if (ENCRYPT)
    radio.encrypt(ENCRYPTKEY);
}

void loop() {
  //Ultrasonic Sensor(Units in centimeters)
  //a=sr04.Distance();

  //Accelerometer Sensor
  //Wire.beginTransmission(MPU_addr);
  //Wire.write(0x3B);  // starting with register 0x3B (ACCEL_XOUT_H)
  //Wire.endTransmission(false);
  //Wire.requestFrom(MPU_addr,14,true);  // request a total of 14 registers
  //AcX=Wire.read()<<8|Wire.read();  // 0x3B (ACCEL_XOUT_H) & 0x3C (ACCEL_XOUT_L)    
  //AcY=Wire.read()<<8|Wire.read();  // 0x3D (ACCEL_YOUT_H) & 0x3E (ACCEL_YOUT_L)
  //AcZ=Wire.read()<<8|Wire.read();  // 0x3F (ACCEL_ZOUT_H) & 0x40 (ACCEL_ZOUT_L)
  //GyX=Wire.read()<<8|Wire.read();  // 0x43 (GYRO_XOUT_H) & 0x44 (GYRO_XOUT_L)
  //GyY=Wire.read()<<8|Wire.read();  // 0x45 (GYRO_YOUT_H) & 0x46 (GYRO_YOUT_L)
  //GyZ=Wire.read()<<8|Wire.read();  // 0x47 (GYRO_ZOUT_H) & 0x48 (GYRO_ZOUT_L)
  //delay(100);

  //Communications Code
  String finalString = "";
  if (radio.receiveDone()) // Got one!
  {
    // The actual message is contained in the DATA array,
    // and is DATALEN bytes in size:
    for(byte i = 0; i < radio.DATALEN; i++)
      finalString += (char)radio.DATA[i];
    int motorPower1 = finalString.substring(0,2).toInt();
    int motorPower2 = finalString.substring(2,4).toInt();
    int motorPower3 = finalString.substring(4,6).toInt();
    int motorPower4 = finalString.substring(6,8).toInt();
    
    //Motor Code
    MotorOne.write(motorPower1);
    MotorTwo.write(motorPower2);
    MotorThree.write(motorPower3);
    MotorFour.write(motorPower4);
    delay(150);   
    
    // Send an ACK if requested.
    // (You don't need this code if you're not using ACKs.)
    if (radio.ACKRequested())
    {
      radio.sendACK();
      Serial.println("ACK sent");
    } 
  }
  
  
}


