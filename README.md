# Supports-input-of-postal-code-and-transfer-code
## 郵便番号と金融機関コードの入力補助フォーム例
**注意：サンプルファイル中の郵便番号と金融機関コードのリストは、ある時点で[こちらのサイト](https://www.post.japanpost.jp/zipcode/download.html)等から独自に入手したものです。「内容に誤りが含まれる可能性がある」「最新情報ではない」等のリスクがあることにご留意ください。**
## 操作方法
入力フォーム画面  
![入力フォームサンプル](https://github.com/takayuki818/Supports-input-of-postal-code-and-transfer-code/assets/147408435/e5229447-7b03-4b92-acda-f96d91bee1b2)
#### 郵便番号入力補助
「郵便番号・住所入力補助」ボタン押下により↓のユーザーフォームが表示されます  
![住所入力・郵便番号自動検索フォーム](https://github.com/takayuki818/Supports-input-of-postal-code-and-transfer-code/assets/147408435/3d4da891-75f2-49df-81bc-1149326004b5)  
赤字はオブジェクト名  
![フォームオブジェクト名](https://github.com/takayuki818/Supports-input-of-postal-code-and-transfer-code/assets/147408435/13dbb489-a586-4c95-b41c-0638a5ece2f9)  
自治体名以降の住所漢字を入力すると、前方一致検索によりドロップダウンリストが表示されます  
![ドロップダウンリスト](https://github.com/takayuki818/Supports-input-of-postal-code-and-transfer-code/assets/147408435/315e4cdc-948f-48df-9bfe-24d2b2077750)  
自治体名-大字名をドロップダウンリストから選択すると、郵便番号が表示されます。  
→「決定・代入」ボタン押下で入力フォームに代入します  
![郵便番号表示](https://github.com/takayuki818/Supports-input-of-postal-code-and-transfer-code/assets/147408435/1a56e355-edb4-4224-96f1-5e0ef2d2a94d)  
#### 金融機関コード自動入力
入力フォーム中で、  
1.「本店名」を入力→本店名コードを自動で検索代入  
2.「支店名」を入力→支店名コードを自動で検索代入  
※支店名コード特定には本店名コードを必要とするため、1→2の順で操作する必要があります。
