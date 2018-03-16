Remove_email_duplicates
Remove email duplicates in MS Outlook.

When changing your email software, what you should do is exporting emails from old software and importing these emails to new software. Outlook accepts .pst file for importing emails from other software. If there are several .pst files and if there are duplicates of the same emails, you must want to remove duplicates.

Outlook offers Visual Basic for Application (VBA). This code is written in VBA.

Motivation to develop this script
Email duplicates occurred when importing emails from old email software to Microsoft Outlook. If the number of emails is small, I would just remove email duplicates manually. However, if there are more than 10,000 emails, the manual method is not realistic. Thus, I wrote up this script.

When I writing this script, I referred the code in the URI shown below. https://outlooklab.wordpress.com/2014/02/08/%E9%87%8D%E8%A4%87%E3%81%97%E3%81%9F%E3%83%A1%E3%83%BC%E3%83%AB%E3%82%92%E5%89%8A%E9%99%A4%E3%81%99%E3%82%8B%E3%83%9E%E3%82%AF%E3%83%AD/

This code does not work in my environment, and this code takes a lot of minutes to execute. Thus, I wanted to execute the code with small computation time.

日本語
このVBAスクリプトは，Microsoft Outlookを使っている際に重複したメールが存在する場合に，その重複を解消してくれるものです。OutlookはVBAが使えるので，VBAで書きました。

このスクリプトは作った動機
メールのソフトウェア変更に伴い，いままで使っていたソフトウェアに含まれているメールをインポートした際に，メールの重複が発生しました。メールが少なければ手動で取り除けばよいのですが，メールがたとえば1万件あった場合には手動での重複解消は非現実的です。そのため，このスクリプトを書きました。このスクリプトを書くにあたり，下記のコードを参考にしました。下記コードは私の環境ではエラーが起きること，また実行時間が非常に長いことから，少ない計算量ですばやく重複を取り除くようにしました。
