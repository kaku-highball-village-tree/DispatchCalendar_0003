# DispatchCalendar_0003

## VBA: B1 から AE1 まで連続日付を入力する

`src/FillDates.bas` には、Excel の B1 に入力された開始日を基準に、B1:AE1 の空白セルへ連続日付を入力する VBA マクロを用意しています。

たとえば B1 に `2026/6/1` が入っている状態で `FillBlankDatesFromB1ToAE1` を実行すると、C1 から AE1 までの空白セルに `2026/6/2` ～ `2026/6/30` が入力されます。

### 使い方

1. Excel で `Alt + F11` を押して VBA エディターを開きます。
2. 標準モジュールを追加するか、`src/FillDates.bas` をインポートします。
3. B1 に開始日を入力します。
4. `FillBlankDatesFromB1ToAE1` を実行します。

### コード例

```vb
Sub 実行例()
    FillBlankDatesFromB1ToAE1
End Sub
```
