# Contributing

Thanks for wanting to help. Soroka is a personal project, so a few things make a contribution easy to take.

## Before you start

- **Open an issue first** for anything bigger than a small fix, so we agree on the idea before you spend time on it.
- **A security problem** goes to a [private report](SECURITY.md), not to an issue or a pull request.

## Build and run

```bash
npm ci
npm start          # run from source
npm run dist       # build the installer into release/
```

- Inside a VS Code terminal, remove `ELECTRON_RUN_AS_NODE` first. VS Code sets it, and with it Electron starts as plain Node.
- Close the installed Soroka before `npm start`. Both use the same single-instance lock and the same port.
- The dev run uses the same data folder, `%APPDATA%\Soroka`, as the installed app. Never commit anything from it.

## A change

- `npx tsc` must pass with no errors.
- Try the change on your own Teams. Use your own chat with yourself («(you)» in the chat list) for anything that sends, and delete the test messages afterwards.
- Every text the user sees goes into `src/i18n.ts` in all three languages: Ukrainian, English, Russian.
- Keep the change to what it is for: no unrelated renames or reformatting.
- Comments say why the code does what it does, not what it does.
- Teams changes its page often. A new selector deserves a comment saying where it was seen, or that it was not seen live yet.

## Safety rules the code keeps

- Nothing is sent to a chat without the user's click.
- What the app asks an AI on its own goes only to an AI without tools.
- Chats and people on «Not to AI» never reach any AI.
- No telemetry, and no server of the project's own.

A change that weakens one of these will not be taken.

## License

By contributing you agree that your contribution is licensed under the [Apache-2.0 license](LICENSE) of this project.
