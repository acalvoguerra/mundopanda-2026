const ALLOWED_ORIGIN = "https://acalvoguerra.github.io";

function corsHeaders(request) {
  const origin = request.headers.get("Origin");
  const allowOrigin = origin === ALLOWED_ORIGIN ? origin : ALLOWED_ORIGIN;

  return {
    "Access-Control-Allow-Origin": allowOrigin,
    "Access-Control-Allow-Methods": "POST, OPTIONS, GET",
    "Access-Control-Allow-Headers": "Content-Type, X-Admin-Key",
    "Access-Control-Max-Age": "86400",
    "Vary": "Origin",
  };
}

function jsonResponse(request, body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: {
      ...corsHeaders(request),
      "Content-Type": "application/json; charset=utf-8",
    },
  });
}

async function githubGetFile(env) {
  const url = `https://api.github.com/repos/${env.GITHUB_OWNER}/${env.GITHUB_REPO}/contents/${env.GITHUB_FILE_PATH || "ranking.xlsx"}?ref=${env.GITHUB_BRANCH || "main"}`;
  const res = await fetch(url, {
    headers: {
      "Authorization": `Bearer ${env.GITHUB_TOKEN}`,
      "Accept": "application/vnd.github+json",
      "User-Agent": "mundopanda-uploader",
    },
  });

  if (res.status === 404) return null;
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`GitHub GET failed ${res.status}: ${text}`);
  }
  return await res.json();
}

async function githubPutFile(env, base64Content, sha) {
  const path = env.GITHUB_FILE_PATH || "ranking.xlsx";
  const url = `https://api.github.com/repos/${env.GITHUB_OWNER}/${env.GITHUB_REPO}/contents/${path}`;
  const payload = {
    message: `Actualizar ${path}`,
    content: base64Content,
    branch: env.GITHUB_BRANCH || "main",
  };
  if (sha) payload.sha = sha;

  const res = await fetch(url, {
    method: "PUT",
    headers: {
      "Authorization": `Bearer ${env.GITHUB_TOKEN}`,
      "Accept": "application/vnd.github+json",
      "Content-Type": "application/json",
      "User-Agent": "mundopanda-uploader",
    },
    body: JSON.stringify(payload),
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`GitHub PUT failed ${res.status}: ${text}`);
  }
  return await res.json();
}

function arrayBufferToBase64(buffer) {
  let binary = "";
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

export default {
  async fetch(request, env) {
    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: corsHeaders(request) });
    }

    if (request.method === "GET") {
      return jsonResponse(request, { ok: true, service: "mundopanda-uploader" });
    }

    if (request.method !== "POST") {
      return jsonResponse(request, { ok: false, error: "Método no permitido" }, 405);
    }

    const adminKey = request.headers.get("X-Admin-Key") || "";
    if (!env.ADMIN_SECRET || adminKey !== env.ADMIN_SECRET) {
      return jsonResponse(request, { ok: false, error: "Clave de administración incorrecta" }, 401);
    }

    const contentType = request.headers.get("Content-Type") || "";
    if (!contentType.includes("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") &&
        !contentType.includes("application/octet-stream")) {
      return jsonResponse(request, { ok: false, error: "El fichero debe ser un .xlsx" }, 400);
    }

    const buffer = await request.arrayBuffer();
    if (!buffer || buffer.byteLength < 1000) {
      return jsonResponse(request, { ok: false, error: "El fichero parece vacío o no válido" }, 400);
    }

    try {
      const current = await githubGetFile(env);
      const base64Content = arrayBufferToBase64(buffer);
      const result = await githubPutFile(env, base64Content, current?.sha);
      return jsonResponse(request, {
        ok: true,
        message: "Excel actualizado correctamente",
        commit: result.commit?.html_url || null,
      });
    } catch (error) {
      return jsonResponse(request, { ok: false, error: error.message }, 500);
    }
  },
};
