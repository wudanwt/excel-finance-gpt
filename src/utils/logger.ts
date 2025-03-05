import { config } from '../config/config';

export class Logger {
  private static instance: Logger;
  private isDebugEnabled: boolean;

  private constructor() {
    this.isDebugEnabled = config.env.isDebug;
  }

  public static getInstance(): Logger {
    if (!Logger.instance) {
      Logger.instance = new Logger();
    }
    return Logger.instance;
  }

  public debug(message: string, ...args: any[]): void {
    if (this.isDebugEnabled) {
      console.debug(`[DEBUG] ${message}`, ...args);
    }
  }

  public info(message: string, ...args: any[]): void {
    console.info(`[INFO] ${message}`, ...args);
  }

  public warn(message: string, ...args: any[]): void {
    console.warn(`[WARN] ${message}`, ...args);
  }

  public error(message: string, error?: unknown, ...args: any[]): void {
    if (error instanceof Error) {
      console.error(`[ERROR] ${message}`, error.message, ...args);
      if (error.stack) {
        console.error(error.stack);
      }
    } else {
      console.error(`[ERROR] ${message}`, error, ...args);
    }
  }
}

export const logger = Logger.getInstance();
