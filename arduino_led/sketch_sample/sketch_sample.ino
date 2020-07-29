
#include <SPI.h>
#include <bitBangedSPI.h>
#include <MAX7219_Dot_Matrix.h>

const byte chips = 4; // ８x８ドットLEDマトリックスの数

// 12 chips (display modules), hardware SPI with load on D10
//MAX7219_Dot_Matrix display (chips, 10, 11, 13);  // Chips / LOAD ,DIN,CLK
MAX7219_Dot_Matrix display (chips,10);  // Chips / LOAD ,DIN,CLK

const char message [] = "        (^^) SYOGAKUSEI PROGRAMMING <^_^>"; // 半角英数字のみ
const char message_yyyy [] = "2021"; // 半角英数字のみ
const char message_mmdd [] = "0321"; // 半角英数字のみ

// ここに起動時に一回実行されるプログラムをかきます
void setup ()
{
  display.begin ();
}  // end of setup

void scrollDisplay ()
{
  int  messageOffset=0;
  while(1){

    display.sendSmooth (message, messageOffset); // 酢黒＾るする文字をセットします
  
    // 文字を表示し終わったら終了します
    if (messageOffset++ >= (int) (strlen (message) * 8)){
      break;
    } 
    delay(50UL);  // スクロールで動く速さ、100ミリ秒でスクロールする
  }
}  // end of updateDisplay

void flashDisplay()
{
  // 徐々に明るくしていきます
  for(int i=0;i<16;i++){
    display.setIntensity(i); // 明るさを調整
    display.sendString(message_yyyy); // 文字を表示
    delay(200UL);  // 明るくする速さ、200msごとに徐々に明るくする
  }
  delay(1000UL);  // 1秒まつ
  display.sendString("    "); // 画面を消す
  delay(1000UL);
  
  for(int i=0;i<16;i++){
    display.setIntensity(i);
    display.sendString(message_mmdd);
    delay(200UL);
  };
  delay(1000UL);
}

// ここに記述したプログラムはループ実行されます
void loop () 
{ 
    scrollDisplay ();  // スクロールする文字
    flashDisplay();    // フラッシュする文字       
}  // end of loop
