const SHORT_HOSTS = new Set(["e.tb.cn", "m.tb.cn", "s.tb.cn", "c.tb.cn"]);

module.exports = async function handler(req, res) {
  if (req.method !== "GET") return res.status(405).json({ error: "仅支持 GET" });
  const raw = Array.isArray(req.query.url) ? req.query.url[0] : req.query.url;
  if (!raw || raw.length > 2048) return res.status(400).json({ error: "缺少或链接过长" });

  let current;
  try {
    current = new URL(raw);
    if (current.protocol !== "https:" || !SHORT_HOSTS.has(current.hostname)) throw new Error();
  } catch {
    return res.status(400).json({ error: "仅支持淘宝官方短链" });
  }

  try {
    for (let i = 0; i < 6; i += 1) {
      if (!SHORT_HOSTS.has(current.hostname)) {
        res.setHeader("Cache-Control", "s-maxage=86400, stale-while-revalidate");
        return res.status(200).json({ url: current.href });
      }

      const response = await fetch(current, {
        redirect: "manual",
        signal: AbortSignal.timeout(8000),
        headers: { "user-agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 Mobile/15E148" },
      });

      const location = response.headers.get("location");
      if (location) {
        current = new URL(location, current);
        continue;
      }

      if (Number(response.headers.get("content-length") || 0) > 1_000_000) throw new Error("响应过大");
      const html = await response.text();
      const match = html.match(/(?:location(?:\.href|\.replace)?\s*[=(]|url\s*=)\s*["']?([^"'<>\s)]+)/i);
      if (!match) throw new Error("短链没有返回目标地址");
      current = new URL(match[1].replace(/&amp;/g, "&"), current);
    }
    throw new Error("跳转次数过多");
  } catch (error) {
    return res.status(502).json({ error: error.message || "短链还原失败" });
  }
};
