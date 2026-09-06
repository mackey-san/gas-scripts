# CLAUDE.md

このファイルは、このリポジトリで Claude Code が作業する際のガイドラインです。

## プロジェクト概要

Google Apps Script (GAS) のスクリプトを管理するリポジトリです。

## Git運用ルール

**コードを変更するたびに、必ず GitHub にプッシュすること。** これはこのプロジェクトの最優先ルールです。

具体的には以下の手順を徹底する。

1. ファイルを変更したら、その都度 `git add` → `git commit` → `git push` を行う（変更をまとめて後回しにしない）
2. コミットメッセージは変更内容が分かる簡潔な日本語または英語で記述する
3. リモートは `origin` (`https://github.com/mackey-san/gas-scripts.git`) の `main` ブランチを使用する
4. push前に `git status` で意図しないファイル（認証情報、`.clasp.json` のシークレット、`node_modules` など）が含まれていないか確認する
5. force push (`git push --force`) は明示的な指示がない限り行わない

## GASプロジェクトに関する注意事項

- `clasp` を利用する場合、`.clasp.json` に含まれる `scriptId` はプロジェクト固有のため、認証トークンなどの秘匿情報をコミットしないよう注意する
- `.clasprc.json` や API キー・認証情報はコミット対象外とする（`.gitignore` で除外する）
