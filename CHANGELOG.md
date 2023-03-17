## [Unreleased]

## [0.3.2](https://github.com/CLOUDs-Inc/ibnet_contract/releases/tag/0.3.2)

### Fixed

- [#201](../../issues/201) 2054/9054 の実金シートで「約定元本」「利用可能額」を空欄にしたい

## [0.3.1](https://github.com/CLOUDs-Inc/ibnet_contract/releases/tag/0.3.1)

### Added

- 出力先ディレクトリやファイル名の整理
- 請求書合計機能
- 貸付金額合計機能
- 商品ごとの金消に連帯保証人を埋め込む機能
- 金消の連帯保証人を適切に追加する処理
- 2054, 9054 に対しての特殊処理の追加
- セットアップスクリプトの追加

### Fixed

- 置換によって画像が消えてしまう問題の修正
- 実金のテーブル参照先が不正だったものを修正
- セルのコピー時に背景色もコピーするよう修正
- xlsx ファイルの置換が失敗してしまう問題の修正

## [0.2](https://github.com/CLOUDs-Inc/ibnet_contract/releases/tag/0.2)

### Added

- 複数物件入力に対応
- 入力シートに複数列の入力がある場合に行ごとに出力をする
- 実金シートのレイアウト変更

### Removed

- 一時的に帳票出力の機能を off

## [0.1](https://github.com/CLOUDs-Inc/ibnet_contract/releases/tag/0.1)

### Added

- 70n の入力から出力までの実装
