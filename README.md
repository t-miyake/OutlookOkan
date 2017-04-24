おかん for Outlook (メール誤送信防止アドイン)
========

おかん for Outlookは、Microsoft Office Outlook用アドインです。  

誤送信を防止するため、メールの送信前に確認ウィンドウを表示します。  
おかんのように色々心配して確認してくれるアドインです。  

機密の関わるメールにおいて、完全なオープンソースのため、安心してご利用いただけます。  
また、キーワードによる警告や、自動CC/BCC追加機能など、便利なオプション機能もあります。  

オープンソースかつ無料でご利用いただけますが、無サポート、無保証です。([ライセンス](https://github.com/t-miyake/OutlookOkan/blob/master/LICENSE))  
専用のカスタマイズやサポートが必要な場合は、個別にご相談ください。  

送信前の確認ウインドウ  
![Screenshot](https://github.com/t-miyake/OutlookOkan/blob/master/Screenshots/Screenshot_v0.9_01.png)  

設定ウィンドウ  
![Screenshot](https://github.com/t-miyake/OutlookOkan/blob/master/Screenshots/Screenshot_v0.9_02.png) 

## 機能一覧(概要)

- メール送信前の確認 
    - メール送信前に確認ウインドウを表示し、全ての項目にチェックしないと送信ができない仕様
    - 社外ドメインは赤文字で表示
    - 件名や添付ファイルの一覧を表示
    - 添付ファイルの添付漏れを警告

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

## 使い方
[Wiki](https://github.com/t-miyake/OutlookOkan/wiki/Manual)に記載します。

## 注意事項
現在、公開しているバージョンは、最低限の動作確認を行ったプレビュー版です。  
既知の不具合や課題は、[Wiki](https://github.com/t-miyake/OutlookOkan/wiki/Known-Issues)に記載しています。  

また、Windows10 Pro Creators Updade (15063) と Office 2016 (32bit版) でのみ、  
動作確認を行っています。
