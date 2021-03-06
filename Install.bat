@cls
@echo **********************************************************
@echo ■ 確認事項
@echo **********************************************************
@echo 現在、Excelを起動している場合は、終了してください。
@pause

@echo **********************************************************
@echo マクロファイルをコピーします..
@cd bin
@copy /Y *.xla "%APPDATA%\Microsoft\Addins"

@echo **********************************************************
@echo ■ アドインを有効化して下さい
@echo **********************************************************
@echo Office2003以前を利用の場合
@echo １　Excelの画面上部 [ツール(T)] → [アドイン]をクリック
@echo ２　アドイン名の左にあるチェックボックスをチェック
@echo ３　[ＯＫ]をクリック
@echo ----------------------------------------------------------
@echo Office2007/2010を利用の場合
@echo 添付の
@echo 【！】インストール手順 Excel2007or2010 を利用の場合.docx
@echo をご覧下さい。
@echo __________________________________________________________

@echo 以上で、インストール作業は終了です。
@pause

