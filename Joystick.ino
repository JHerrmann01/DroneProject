const int SW_pin = 2;
const int X_pin = 0;
const int Y_pin = 1;

void setup()
{
  pinMode(SW_pin, INPUT);
  digitalWrite(SW_pin, HIGH);
  pinMode(X_pin, INPUT);
  pinMode(Y_pin, INPUT);
  Serial.begin(9600);
}

void loop()
{
  //When encoding the values, just send finalValue
  
  int xVal = analogRead(X_pin);
  int yVal = analogRead(Y_pin);
  
  //X Value
  if((486 < xVal) && (xVal < 536))
  {
    Serial.print("X-Axis is Stationary : 0%");
  }
  else if(xVal >= 536)
  {
    int finalValue = (.205338809)*(xVal - 536);
    String finalString = String(finalValue);
    Serial.print("X-Axis is directed backwards : " + finalString + "%");    
  }
  else if(xVal <= 486)
  {
    int finalValue = (-.2057613169)*(486-xVal);
    String finalString = String(finalValue);
    Serial.print("X-Axis is directed forwards : " + finalString + "%");  }

  Serial.print("\n");

  //Y Value
  if((486 < yVal) && (yVal < 536))
  {
    Serial.print("Y-Axis is Stationary : 0%");
  }
  else if(yVal >= 536)
  {
    int finalValue = (.205338809)*(yVal - 536);
    String finalString = String(finalValue);
    Serial.print("Y-Axis is directed left : " + finalString + "%");
  }
  else if(yVal <= 486)
  {
    int finalValue = (-.2057613169)*(486-yVal);
    String finalString = String(finalValue);
    Serial.print("Y-Axis is directed right : " + finalString + "%");
  }
  Serial.print("\n\n");
  delay(500);
  
}

