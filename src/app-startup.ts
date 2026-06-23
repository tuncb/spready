export interface MainStartupOptions {
  consoleOutputFilePath?: string;
  help: boolean;
  workbookFilePath?: string;
}

const WORKBOOK_FILE_EXTENSION = ".spready";

function getArgumentName(token: string) {
  const separatorIndex = token.indexOf("=");

  return token.slice(2, separatorIndex === -1 ? undefined : separatorIndex);
}

function getInlineArgumentValue(token: string) {
  const separatorIndex = token.indexOf("=");

  if (separatorIndex === -1) {
    return undefined;
  }

  return token.slice(separatorIndex + 1);
}

function readArgumentValue(argv: string[], index: number, token: string) {
  const inlineValue = getInlineArgumentValue(token);

  if (inlineValue !== undefined) {
    if (inlineValue.length === 0) {
      throw new Error(`Missing value for --${getArgumentName(token)}.`);
    }

    return {
      nextIndex: index,
      value: inlineValue,
    };
  }

  const value = argv[index + 1];

  if (!value || value.startsWith("--")) {
    throw new Error(`Missing value for --${getArgumentName(token)}.`);
  }

  return {
    nextIndex: index + 1,
    value,
  };
}

export function parseMainStartupOptions(argv: string[]): MainStartupOptions {
  const options: MainStartupOptions = {
    help: false,
  };

  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];

    if (token === "-h") {
      options.help = true;
      continue;
    }

    if (!token.startsWith("--")) {
      if (
        options.workbookFilePath === undefined &&
        token.toLowerCase().endsWith(WORKBOOK_FILE_EXTENSION)
      ) {
        options.workbookFilePath = token;
      }

      continue;
    }

    const name = getArgumentName(token);

    switch (name) {
      case "help":
        if (getInlineArgumentValue(token) !== undefined) {
          throw new Error(`--${name} does not accept a value.`);
        }

        options.help = true;
        break;
      case "console-output": {
        const parsed = readArgumentValue(argv, index, token);

        options.consoleOutputFilePath = parsed.value;
        index = parsed.nextIndex;
        break;
      }
      default:
        break;
    }
  }

  return options;
}

export function getMainHelpText(commandName = "spready") {
  return [
    "Spready",
    "",
    `Usage: ${commandName} [options]`,
    "",
    "Options:",
    "  -h, --help                         Show this help message.",
    "      --console-output FILE          Print a .spready workbook to stdout and exit.",
    "",
    "Arguments:",
    "  FILE                               Open a .spready workbook.",
  ].join("\n");
}
