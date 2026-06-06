# 変数命名規則 (Variable Naming Rules)

## 目的

本ドキュメントは、このプロジェクトで使用する変数命名規則を定義します。

この規則の目的は次のとおりです。

- 可読性の向上
- 型を即座に判別できるようにする
- デバッグを容易にする
- AI支援コーディングの精度を高める
- 初学者がコード構造を理解しやすくする
- 不明瞭な略語を防ぐ

---

# 基本規則

## 明示的な型宣言（必須）

すべての変数は、**必ず明示的な型宣言**を使用しなければなりません。  
`auto` などによる暗黙的・推論的な宣言は使用しないでください。

悪い例:

```cpp
auto x = 10;
```

良い例:

```cpp
int iItemCount = 10;
```

---

# プレフィックス規則（必須）

変数名は原則として、型に応じて以下のプレフィックスを必ず使用してください。

## int

`int` 型の変数は、`i` で始めること。

```cpp
int iItemCount;
int iCoordinateX;
int iCoordinateY;
int iCoordinateZ;
```

## unsigned int

`unsigned int` 型の変数は、`ui` で始めること。

```cpp
unsigned int uiIndex;
unsigned int uiTotalCount;
```

## char

`char` 型の変数は、`c` で始めること。

```cpp
char cCharacter;
char cSeparator;
```

## unsigned char

`unsigned char` 型の変数は、`uc` で始めること。

```cpp
unsigned char ucBuffer;
unsigned char ucCharacterCode;
```

## long

`long` 型の変数は、`l` で始めること。

```cpp
long lFileSize;
long lElapsedTime;
```

## BOOL

`BOOL` 型の変数は、`b` で始めること。

```cpp
BOOL bIsVisible;
BOOL bIsSuccess;
BOOL bIsCompleted;
```

## float

`float` 型の変数は、`f` で始めること。

```cpp
float fTemperature;
float fScaleValue;
```

## double

`double` 型の変数は、`d` で始めること。

```cpp
double dAverageValue;
double dCalculationResult;
```

## 文字列型

文字列型の変数は、`psz` で始めること。

```cpp
char* pszTextFile;
char* pszFilePath;
char* pszMessageText;
```

## 構造体オブジェクト

構造体オブジェクトは、`obj` で始めること。

```cpp
PERSON objPerson;
FILE_INFO objFileInfo;
```

## 構造体ポインタ

構造体ポインタは、`p` で始めること。

```cpp
PERSON* pPerson;
FILE_INFO* pFileInfo;
```

---

# 構造体メンバー規則

構造体内の変数名は、**入門段階の方針として通常の変数と同じ規則**を適用してください。  
（= 型に応じた同一プレフィックス規則を守る）

例:

```cpp
typedef struct tagPERSON
{
    int iAge;
    char* pszPersonName;
    BOOL bIsActive;

} PERSON;
```

---

# 短すぎる名前の禁止（必須）

意味が伝わりにくい短すぎる変数名は禁止します。  
特に `iX`, `iY`, `iZ` のような最小文字数の命名は避け、役割を含む語を使用してください。

悪い例:

```cpp
int iX;
int iY;
int iZ;
char* pszTf;
BOOL bFlg;
```

良い例:

```cpp
int iCoordinateX;
int iCoordinateY;
int iCoordinateZ;
char* pszTextFile;
BOOL bIsVisible;
```

---

# 不明瞭な略語の禁止（必須）

ルール上プレフィックスが正しくても、**不明瞭な省略語は使用しない**でください。  
例: `pszTf` は禁止し、`pszTextFile` のように意味が明確な語を使ってください。

悪い例:

```cpp
pszTmp
pszBuf
pszTxt
```

良い例:

```cpp
pszTemporaryFilePath
pszReceiveBuffer
pszTextFilePath
```

---

# 命名の明確性に関する補足ルール（追加反映）

以下を必須とします。

1. 型プレフィックスの後ろには、**用途が分かる語**を続けること。  
2. 必要以上の省略をしないこと（`Tf`, `Flg`, `Buf` など曖昧な略語を避ける）。  
3. 座標・位置などでも、`iX` ではなく `iCoordinateX` のように**意味語を保持**すること。  
4. 同一スコープ内で似た名前を乱立させず、読み間違いを防ぐこと。  
5. 初学者が見て意味を推測できる語を優先すること。

---

# AI支援コーディング

これらの命名規則は、次の改善を目的としています。

- AIによるコード生成
- AIによるコード補完
- AIによるコード可読性
- AI支援デバッグ

明確な変数名は、以下のようなツールの精度向上に寄与します。

- ChatGPT
- Claude Code
- GitHub Copilot
- Cursor

---

# まとめ

このプロジェクトは、以下を重視します。

- 可読性
- 保守性
- 初学者への親和性
- 明示的な型
- 型と一致したプレフィックス
- 不明瞭な略語の排除
- 短すぎる命名の禁止
- AIとの親和性
