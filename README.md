# エクセルを用いたパラメータ化テスト

このサンプルは、エクセルファイルを使用してパラメータ化テストを行うものです。

`SampleTest`内のテストを起動すると、エクセルファイルに指定した日付を用いて起動したサイトの宿泊日を入力します。



## 参考にしたサイト

1. [shimashima35](https://github.com/shimashima35/codezine-sample)さんが提供している Selenideのサンプルコード
   Selenideを使用して、[日本Seleniumユーザコミュニティで提供されているテスト用サイト](http://www.selenium.jp/test-site)向けのテストを作成している。このコードがベースです。

2. [YoshikiIto](https://github.com/YoshikiIto/selenium-junit4-datadriven-sample)さんが提供しているエクセルを使用したパラメータ化テストのサンプル
   エクセルからデータを取得する仕組みは、このサイトで提供されているものをまるっと活用しています



## 使用している技術等

- JUnit4 + Eclipse + Maven
- Selenide
- POI