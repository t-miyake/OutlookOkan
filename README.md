おかん for Outlook (メール誤送信防止アドイン)
========

おかん for Outlookは、Microsoft Office Outlook用アドインです。  

誤送信を防止するため、メールの送信前に確認ウィンドウを表示します。  
おかんのように色々心配して確認してくれるアドインです。  

機密の関わるメールにおいて、完全なオープンソースのため、安心してご利用いただけます。  
また、キーワードによる警告や、自動CC/BCC追加機能など、便利なオプション機能もあります。  

ダウンロードは[releases](https://github.com/t-miyake/OutlookOkan/releases)からできます。

オープンソースかつ無料でご利用いただけますが、無サポート、無保証です。([ライセンス](https://github.com/t-miyake/OutlookOkan/blob/master/LICENSE))  
専用のカスタマイズやサポートが必要な場合は、個別にご相談ください。  

送信前の確認ウインドウ  
![Screenshot](https://github.com/t-miyake/OutlookOkan/blob/master/Screenshots/Screenshot_v2.0.0_01.png)  

設定ウィンドウ  
![Screenshot](https://github.com/t-miyake/OutlookOkan/blob/master/Screenshots/Screenshot_v2.0.0_02.png) 

## 機能一覧(概要)

- メール送信前の確認 
    - メール送信前に確認ウインドウを表示し、全ての項目にチェックしないと送信ができない仕様
    -  同一ドメインへのメールなど、送信前の確認を表示しないことも可能
    - 社外ドメインは赤文字で表示
    - 件名や送信者のアドレス、添付ファイルの一覧、メール本文を表示
    - 添付ファイルの添付漏れや大容量の添付ファイルを警告

- 送信禁止機能
    - 指定した宛先やドメインへのメール送信を禁止
    - 指定したキーワードが本文に含まれるメールの送信を禁止
    - 送信禁止に該当する場合、禁止の旨とその理由を表示

- メール送信前の確認 
    - メール送信前に確認ウインドウを表示し、全ての項目にチェックしないと送信ができない仕様
    - 社外ドメインは赤文字で表示

- ホワイトリスト
    - ホワイトリストに登録したドメインやアドレスは、確認画面での項目チェックが不要

- 名称と送付先の登録と警告
    - メール本文中に登場する名称と、送付先のアドレスやドメインが一致しない場合、警告を表示

- 警告キーワードの登録と警告
    - 登録したキーワードがメール本文に含まれる場合、登録した警告文を表示

- 警告アドレスの登録と警告
    - 登録したアドレスやドメインへメールを送信する際に、警告文を表示

- 自動CC/BCC追加(キーワード)
    - 指定したキーワードがメール本文に含まれる場合、指定したアドレスを自動でCCやBCCに追加

- 自動CC/BCC追加(宛先)
    - 指定した宛先へのメールに、指定したアドレスを自動でCCやBCCに追加

- 設定のインポート/エクスポート
    - 設定内容をCSVファイルでインポート/エクスポート

- 多言語対応
    - 日本語と英語に対応しており、言語の追加が可能な設計

## 使い方
[Wiki(Manual)](https://github.com/t-miyake/OutlookOkan/wiki/Manual) に記載します。

## 既知の不具合や課題
[Wiki(Known Issues)](https://github.com/t-miyake/OutlookOkan/wiki/Known-Issues) に記載します。

## ロードマップ
[Wiki(Roadmap)](https://github.com/t-miyake/OutlookOkan/wiki/Roadmap) に記載します。
