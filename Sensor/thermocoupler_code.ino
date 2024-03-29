/*************************************************** 
  This is an example for the Adafruit Thermocouple Sensor w/MAX31855K

  Designed specifically to work with the Adafruit Thermocouple Sensor
  ----> https://www.adafruit.com/products/269
  
  These displays use SPI to communicate, 3 pins are required to  
  interface
  Adafruit invests time and resources providing this open source code, 
  please support Adafruit and open-source hardware by purchasing 
  products from Adafruit!
  
  Written by Limor Fried/Ladyada for Adafruit Industries.  
  BSD license, all text above must be included in any redistribution
 ****************************************************/

#include <SPI.h>
#include <Adafruit_MAX31855.h>
#include <SoftwareSerial.h>

// Default connection is using software SPI, but comment and uncomment one of
// the two examples below to switch between software SPI and hardware SPI:

// Example creating a thermocouple instance with software SPI on any three
// digital IO pins.
#define MAXDO   5
#define MAXCS   4
#define MAXCLK  3

#define MAXDO2   11
#define MAXCS2   10
#define MAXCLK2  9

// initialize the Thermocouple
Adafruit_MAX31855 thermocouple(MAXCLK, MAXCS, MAXDO);

Adafruit_MAX31855 thermocouple2(MAXCLK2, MAXCS2, MAXDO2);

// Example creating a thermocouple instance with hardware SPI
// on a given CS pin.
//#define MAXCS   10
//Adafruit_MAX31855 thermocouple(MAXCS);  

void setup() {
  Serial.begin(9600);
 
  while (!Serial) delay(1); // wait for Serial on Leonardo/Zero, etc

  Serial.println("MAX31855 test");
    // wait for MAX chip to stabilize 
  delay(500);

  Serial.println("MAX31855 test");
  delay(5000);
}
void loop() {
  // basic readout test, just print the current temp
  

   double c = thermocouple.readCelsius();
   double c2 = thermocouple2.readCelsius();
   //char text [40];  
   
   if (isnan(c) && isnan(c)) {
     Serial.println("Something wrong with thermocouple!");
     
   } else {
    
     //sprintf(text,"%d,%d\n",((int)c),((int)c2));
     //Serial.println(text); 
     
     Serial.println("C = " + (String)c + "\tC2 = " + (String)c2);
   }
   //Serial.print("F = ");
   //Serial.println(thermocouple.readFarenheit());
 
   delay(1000);
}
