interface Config {
  app: {
    name: string;
    version: string;
  };
  api: {
    openai: {
      baseUrl: string;
      key: string;
      timeout: number;
      maxTokens: number;
      model: string;
      defaultHeaders?: Record<string, string>;
    };
  };
  env: {
    isDevelopment: boolean;
    isDebug: boolean;
  };
}

export const config: Config = {
  app: {
    name: "Excel Finance GPT",
    version: "1.0.0"
  },
  api: {
    openai: {
      baseUrl: process.env.OPENAI_API_URL || "https://api.openai.com/v1",
      key: process.env.OPENAI_API_KEY || "",
      timeout: parseInt(process.env.API_TIMEOUT || "30000"),
      maxTokens: parseInt(process.env.MAX_TOKENS || "1000"),
      model: process.env.OPENAI_API_MODEL || "text-davinci-003",
      defaultHeaders: {
        "Content-Type": "application/json",
      }
    }
  },
  env: {
    isDevelopment: process.env.NODE_ENV === "development",
    isDebug: process.env.DEBUG === "true"
  }
};

if (!config.api.openai.key) {
  console.warn("OpenAI API key is not configured. Please set OPENAI_API_KEY environment variable.");
}

if (!config.api.openai.baseUrl) {
  console.warn("OpenAI API URL is not configured. Using default OpenAI endpoint.");
}
