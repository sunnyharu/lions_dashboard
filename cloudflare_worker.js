/**
 * Cloudflare Worker: 특이사항 업데이트 프록시
 *
 * Worker 환경변수 (Cloudflare 대시보드에서 설정):
 *   GITHUB_TOKEN : GitHub Personal Access Token (workflow 권한)
 */

export default {
  async fetch(request, env) {
    // CORS 헤더
    const corsHeaders = {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
    };

    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: corsHeaders });
    }

    if (request.method !== 'POST') {
      return new Response('Method Not Allowed', { status: 405, headers: corsHeaders });
    }

    let body;
    try {
      body = await request.json();
    } catch {
      return new Response('Invalid JSON', { status: 400, headers: corsHeaders });
    }

    const { date, note } = body;
    if (!date || !note) {
      return new Response('date and note are required', { status: 400, headers: corsHeaders });
    }

    // GitHub Actions workflow_dispatch 트리거
    const githubRes = await fetch(
      'https://api.github.com/repos/sunnyharu/lions_dashboard/actions/workflows/update_note.yml/dispatches',
      {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${env.GITHUB_TOKEN}`,
          'Accept': 'application/vnd.github.v3+json',
          'Content-Type': 'application/json',
          'User-Agent': 'lions-dashboard-worker',
        },
        body: JSON.stringify({
          ref: 'main',
          inputs: { date, note },
        }),
      }
    );

    if (githubRes.status === 204) {
      return new Response(JSON.stringify({ ok: true }), {
        status: 200,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      });
    } else {
      const errText = await githubRes.text();
      return new Response(`GitHub API 오류: ${errText}`, {
        status: 500,
        headers: corsHeaders,
      });
    }
  },
};
