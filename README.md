# Supports-input-of-postal-code-and-transfer-code
## 郵便番号と金融機関コードの入力補助フォーム例
**注意：サンプルファイル中の郵便番号と金融機関コードのリストは、ある時点で[こちらのサイト](https://www.post.japanpost.jp/zipcode/download.html)等から独自に入手したものです。「内容に誤りが含まれる可能性がある」「最新情報ではない」等のリスクがあることにご留意ください。**
## 操作方法
入力フォーム画面  
![入力フォーム画面](https://github.com/takayuki818/Supports-input-of-postal-code-and-transfer-code/assets/147408435/2badc5e3-df93-4c25-9551-16c8456a3d74)  
#### 郵便番号入力補助
「郵便番号・住所入力補助」ボタン押下（または「住所上段」セルの選択）により入力補助フォームが表示されます。  
![UserForm1](https://github.com/takayuki818/Supports-input-of-postal-code-and-transfer-code/assets/147408435/3ed9aeda-8f9b-441d-aa6b-badaf272a39c)
自治体名-大字名の一部を入力すると、中間一致検索フィルタ処理により、  
合致するデータが格納されたリストボックスが表示されます。（サンプル対応範囲：千葉県内）  
![リストボックス展開](https://github.com/takayuki818/Supports-input-of-postal-code-and-transfer-code/assets/147408435/53330b7c-b924-4491-8762-7bb86cbfafe7)  
赤字はオブジェクト名  
![オブジェクト配置](https://github.com/takayuki818/Supports-input-of-postal-code-and-transfer-code/assets/147408435/03baa3e9-2601-4398-90be-896bc50402fb)  
リストから対象を選択し、「決定」ボタン押下で郵便番号・住所を入力フォームに代入します。  
#### 金融機関コード自動入力
「振込先入力補助」ボタン押下（または「本店名」セルの選択）により、同様のユーザーフォームが表示されます。  
※郵便番号入力補助と同様の操作感となります。
![UserForm2](https://github.com/takayuki818/Supports-input-of-postal-code-and-transfer-code/assets/147408435/ac65fbcd-6b4d-4393-9944-253bf7409a20)  
1.「本店名」を入力　　　→　本店名リストを展開  
2.本店名リストから選択　→　支店名リストを展開  
3.「支店名」を入力　　　→　本店コードと合致する支店名から絞り込み  
※支店の特定には本店情報を必要とするため、先に本店名を特定（リスト選択）する必要が必要があります。
