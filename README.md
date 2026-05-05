# organization-scaner

CLI-проект на TypeScript для чтения Excel-файлов и дальнейшей асинхронной обработки данных через Puppeteer.

## Установка

```bash
npm install
```

## Команды

```bash
npm run build
npm start -- "D:\VS Code\cadastr\1.xlsx" xlsx 4
npm run dev -- "D:\VS Code\cadastr\1.xlsx" xlsx 4
```

После сборки CLI можно запустить напрямую:

```bash
node dist/index.js "D:\VS Code\cadastr\1.xlsx" xlsx 4
```

Формат аргументов: `<путь-к-файлу.xlsx> <выходная-директория> <номер-колонки>`.

`--column` использует нумерацию с 1: `1` означает колонку A, `2` означает колонку B.

Приложение читает значения из указанной колонки первого листа Excel-файла и передаёт каждое непустое значение в асинхронный метод `parse`.

## Строгая типизация

TypeScript настроен в strict-режиме через `tsconfig.json`.
