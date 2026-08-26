import worker from "../dist/server/index.js";

type VercelRequest = {
  method?: string;
  url?: string;
  headers: Record<string, string | string[] | undefined>;
  body?: unknown;
};

type VercelResponse = {
  status: (code: number) => VercelResponse;
  setHeader: (name: string, value: string) => void;
  send: (body: Buffer | string) => void;
};

export default async function handler(req: VercelRequest, res: VercelResponse) {
  const forwardedProto = req.headers["x-forwarded-proto"];
  const protocol = Array.isArray(forwardedProto) ? forwardedProto[0] : forwardedProto || "https";
  const forwardedHost = req.headers["x-forwarded-host"] || req.headers.host || "localhost";
  const host = Array.isArray(forwardedHost) ? forwardedHost[0] : forwardedHost;
  const requestUrl = new URL(req.url || "/", `${protocol}://${host}`);
  const headers = new Headers();

  for (const [name, value] of Object.entries(req.headers)) {
    if (value !== undefined) headers.set(name, Array.isArray(value) ? value.join(", ") : value);
  }

  const method = (req.method || "GET").toUpperCase();
  const request = new Request(requestUrl, {
    method,
    headers,
    body: method === "GET" || method === "HEAD" ? undefined : JSON.stringify(req.body ?? {}),
  });
  const response = await worker.fetch(request, {}, {});

  response.headers.forEach((value: string, name: string) => res.setHeader(name, value));
  res.status(response.status).send(Buffer.from(await response.arrayBuffer()));
}
