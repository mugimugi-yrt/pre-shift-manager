# シフト調整用シート作成 / pre-shift sheet maker
author: mugi(Japanese)

## このプロジェクトの狙い
自分が所属している大学機関の中で、シフト調整を行うとき、wordで提出し、シフト調整担当が紙に書き起こして調整をしていた。しかし、それでは書き間違えや提出忘れなどのミスが多発しており、その人為的ミスを減らすため、「シフト調整のためのシフト表」を作ることにした。  
その期間ではGoogle Driveをよく使用していることから、シフト提出をGoogleフォーム、表作成をフォーム回答が集まるGoogleスプレッドシート上でやることにした。スプレッドシート上で表を表示するため、使用言語はGoogle Apps Scriptとした。

## このプログラムの使用方法
ここにあるコードをフォーム回答が収拾されているスプレッドシートのApps Script内に格納する。  
この時、フォームで集めている質問や、作る表の形によって書き込む場所が異なるので、適宜修正する必要がある。  
自分の使用用途として作成するシフト表は、1~3年生が2限,L限,3限,4限,5限勤務する上でどこに入りたいかの調査だったので、以下のような形になっている。  
|     |     | 月曜日 | 火曜日 | 水曜日 | 木曜日 | 金曜日 | 
| --- | --- | -----  | ----- | -----  | ----- | ----- |
|     | 3年 |        |       |        |       |       |
| 2限 | 2年 |        |       |        |       |       |
|     | 1年 |        |       |        |       |       |
|     | 3年 |        |       |        |       |       |
| L限 | 2年 |        |       |        |       |       |
|     | 1年 |        |       |        |       |       |
|     | 3年 |        |       |        |       |       |
| 3限 | 2年 |        |       |        |       |       |
|     | 1年 |        |       |        |       |       |
|     | 3年 |        |       |        |       |       |
| 4限 | 2年 |        |       |        |       |       |
|     | 1年 |        |       |        |       |       |
|     | 3年 |        |       |        |       |       |
| 5限 | 2年 |        |       |        |       |       |
|     | 1年 |        |       |        |       |       |

この表を基に作成しているため、必要に応じて書き込み場所の調整をしてほしい。  
また、フォームの質問は以下のとおりである。

- 学年  
  ->1~4をラジオボタンで選ぶ形式
- 氏名(苗字)  
  ->テキスト入力
- 氏名(名前)  
  ->テキスト入力
- 同じ苗字がいるか？  
  ->yes/noのラジオボタンで選ぶ形式　yesを選ぶと表示名が「苗字(名前1文字目)」のような見え方になる
- シフト希望  
  ->表から勤務できる時間をチェックボックスにチェックを入れて選ぶ形式
- 備考  
  ->テキスト入力

質問内容を変える時、formResponseの番号を変えたりしなければいけないので注意。また、質問を変えたとき、スプレッドシートの方での解答回収が変な順序になってしまうので、変えたらフォーム収集のリンクを解除し、再度リンクしなおすと良い。

## 現在の進捗
質問に答えたら、人の名前が正確に入力できるところまでは確認済みである。  
今後は、どんな機能を追加していくのかシフト調整担当や同僚に聞いて微細なものから大きなものまで機能拡張していく。

## 使用に当たって
このスプレッドシートとフォームを使ったシフト調整表作成システムは、私が所属している大学機関のようにお金をかけて有料システムを使えないことから作成したシステムである。  
このコードは好きに利用していただいて構わないが、こちらに修正願を出されても対応できないのでバグが発生した場合は生成系AIや周りのプログラムに精通している人に解消してもらってほしい。
