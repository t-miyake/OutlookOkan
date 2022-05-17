おかん for Outlook (メール誤送信防止アドイン)
========

English readme is [here](https://github.com/t-miyake/OutlookOkan/blob/master/README_en.md).

おかん for Outlook (Outlook Okan)は、Microsoft Office Outlook用アドインです。  

誤送信を防止するため、メールの送信前に確認ウィンドウを表示します。  
おかんのように色々心配して確認してくれるアドインです。  

機密の関わるメールにおいて、完全なオープンソースのため、安心してご利用いただけます。  
また、キーワードによる警告や、自動Cc/Bcc追加機能など、便利なオプション機能もあります。  

ダウンロードは[releases](https://github.com/t-miyake/OutlookOkan/releases)からできます。  
※アドイン名を無難なものにしたバージョンも併せて配布しています。

オープンソースかつ無料でご利用いただけますが、無サポート、無保証です。([ライセンス](https://github.com/t-miyake/OutlookOkan/blob/master/LICENSE))  
専用のカスタマイズやサポートが必要な場合は、個別にご相談ください。  

送信前の確認ウインドウ  
![Screenshot 1](https://github.com/t-miyake/OutlookOkan/blob/master/Screenshots/Screenshot_v2.5.0_01.png)  

設定ウィンドウ(一般設定)  
![Screenshot 2](https://github.com/t-miyake/OutlookOkan/blob/master/Screenshots/Screenshot_v2.7.0_04.png)

設定ウィンドウ(遅延送信)  
![Screenshot 3](https://github.com/t-miyake/OutlookOkan/blob/master/Screenshots/Screenshot_v2.7.0_05.png)

送信禁止通知  
![Screenshot 4](https://github.com/t-miyake/OutlookOkan/blob/master/Screenshots/Screenshot_v2.5.0_03.png)

バージョン情報  
![Screenshot 5](https://github.com/t-miyake/OutlookOkan/blob/master/Screenshots/Screenshot_v2.6.1_02.png)

## 対応環境

- Windows 7 / 8 / 8.1 / 10 / 11
- Microsoft Outlook 2013 / 2016 / 2019 / 2021 / Microsoft 365 Apps (32bit版 及び 64bit版)
- .NET Framework 4.6.2 以上

## 機能一覧(概要)

- メール送信前の確認など
  - メール送信前に確認ウインドウを表示し、全ての項目にチェックしないと送信ができない仕様
  - 内部(社内)ドメインへのメールなど、送信前の確認を表示しない設定も可能
  - 外部(社外)ドメインは赤文字で表示
  - 件名や送信者のアドレス、添付ファイルの一覧、メール本文を表示
  - 添付ファイルの添付漏れや大容量の添付ファイルを警告
  - 配布リストや連絡先グループを展開して、各宛先を表示 (オン/オフ可)
  - 宛先をドメイン別に並べ替えて表示 (オン/オフ可)
  - 送信元アドレスを常に自動でCcやBccに追加  (オン/オフ可)

- 送信禁止機能
  - 指定した宛先やドメインへのメール送信を禁止
  - 指定したキーワードが本文に含まれるメールの送信を禁止
  - 指定した宛先やドメインへの添付ファイル付きメールの送信を禁止
  - 添付ファイル付きメールの送信を禁止 (オン/オフ可)
  - 連絡先に登録されていない宛先へのメール送信を禁止 (オン/オフ可)
  - 宛先(To/Cc)外部ドメイン数が多い場合に、メールの送信を禁止 (オン/オフ可)
  - 暗号化ZIPファイルが添付されいている場合に、メールの送信を禁止 (オン/オフ可)
  - 送信禁止に該当する場合、禁止の旨とその理由を表示

- 許可リスト
  - 許可リストに登録したドメインやアドレスは、確認画面での項目チェックが不要に

- 名称と送付先の登録と警告
  - メール本文中に登場する名称と、送付先のアドレスやドメインが一致しない場合、警告を表示

- 警告キーワードの登録と警告
  - 登録したキーワードがメール本文や件名に含まれる場合、登録した警告文を表示
  - 常に登録した警告メッセージを表示することも可能

- 警告アドレスの登録と警告
  - 登録したアドレスやドメインへメールを送信する際に、警告文を表示  
  - 警告文を宛先別に設定することも可能

- 宛先(To/Cc)の外部ドメイン数警告とBccへの自動変換
  - 宛先(To/Cc)外部ドメイン数が多い場合の、警告表示
  - 宛先(To/Cc)外部ドメイン数が多い場合の、宛先(To/Cc)外部アドレスのBccへの自動変換
  - 強制的に全ての宛先をBccに変換

- 自動Cc/Bcc追加(キーワード)
  - 指定したキーワードがメール本文に含まれる場合、指定したアドレスを自動でCcやBccに追加

- 自動Cc/Bcc追加(宛先)
  - 指定した宛先へのメールに、指定したアドレスを自動でCCやBccに追加

- 自動Cc/Bcc追加(添付ファイル)
  - ファイルが添付されたメールに、指定したアドレスを自動でCcやBccに追加

- 送信遅延(送信保留)
  - 設定した時間(分単位)だけ、メールの送信を遅延(保留)
  - ドメインやメールアドレス毎に、デフォルトの遅延時間を設定可能

- 添付ファイル名と宛先の紐づけ
  - 添付ファイル名と宛先メールアドレスやドメインを紐づけ、該当しない場合、警告を表示

- 添付ファイルがある場合の宛先ごとの警告
  - 宛先(アドレスやドメイン)ごとに、添付ファイルがある場合の警告文の設定が可能

- メール本文への文言の自動追加
  - メール本文の文頭や末尾に、指定した文言を自動追加可能。

- その他
  - 暗号化ZIPファイルが添付されている場合に警告を表示  (オン/オフ可)

- 設定のインポート/エクスポート
  - 設定内容をCSVファイルでインポート/エクスポート

- 多言語対応
  - 日本語や英語など計10言語に対応しており、言語の追加が可能な設計

## 使い方

[Wiki(Manual)](https://github.com/t-miyake/OutlookOkan/wiki/Manual) に記載します。

## 既知の不具合

[Wiki(Known Issues)](https://github.com/t-miyake/OutlookOkan/wiki/Known-Issues) に記載します。

## ロードマップ

[Wiki(Roadmap)](https://github.com/t-miyake/OutlookOkan/wiki/Roadmap) に記載します。
