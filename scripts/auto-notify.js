// scripts/auto-notify.js
// GitHub Actions에서 매일 자동 실행되는 기한지연 메일 발송 스크립트

const SUPABASE_URL  = (process.env.VITE_SUPABASE_URL || '').replace(/\/$/, '');
const SUPABASE_BASE = SUPABASE_URL.startsWith('http') ? SUPABASE_URL : `https://${SUPABASE_URL}`;
const SUPABASE_ANON = process.env.VITE_SUPABASE_ANON;
const today = new Date().toISOString().slice(0, 10);

function fmtDate(d) {
  if (!d) return '-';
  const dt = new Date(d);
  return `${dt.getFullYear()}.${String(dt.getMonth()+1).padStart(2,'0')}.${String(dt.getDate()).padStart(2,'0')}`;
}

async function supabaseGet(query) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/rentals?${query}`, {
    headers: {
      'apikey': SUPABASE_ANON,
      'Authorization': `Bearer ${SUPABASE_ANON}`,
      'Content-Type': 'application/json',
    }
  });
  if (!res.ok) throw new Error(`Supabase GET 실패: ${await res.text()}`);
  return res.json();
}

async function supabasePatch(id, body) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/rentals?id=eq.${id}`, {
    method: 'PATCH',
    headers: {
      'apikey': SUPABASE_ANON,
      'Authorization': `Bearer ${SUPABASE_ANON}`,
      'Content-Type': 'application/json',
      'Prefer': 'return=minimal',
    },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`Supabase PATCH 실패: ${await res.text()}`);
}

async function sendEmail(item, message) {
  const res = await fetch('https://api.emailjs.com/api/v1.0/email/send', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      service_id:  process.env.EMAILJS_SERVICE_ID,
      template_id: process.env.EMAILJS_TEMPLATE_ID,
      user_id:     process.env.EMAILJS_PUBLIC_KEY,
      accessToken: process.env.EMAILJS_PRIVATE_KEY,
      template_params: {
        to_email:  item.requester_email,
        to_name:   item.requester,
        category:  item.category,
        item_name: item.item_name,
        due_date:  fmtDate(item.due_date),
        purpose:   item.purpose || '-',
        message,
        from_name: 'Bimatrix 인사기획팀',
      }
    })
  });
  if (!res.ok) throw new Error(await res.text());
}

function buildMessage(item) {
  const rank = item.requester_rank ? ' ' + item.requester_rank : '';
  if (item.category === '메일계정') {
    return (
      `${item.requester}${rank}님, 안녕하세요. 인사기획팀입니다.\n\n` +
      `이전에 신청하셨던 ${item.borrower} 메일 계정[${item.item_name}]의 사용 기한[${fmtDate(item.due_date)}]이 지나 연락드립니다.\n\n` +
      `따라서 계정 비활성화 처리를 하고자 하는데 진행해도 되는지에 대하여 회신 부탁드립니다.\n\n` +
      `만약 연장이 필요할 경우, 연장 사유와 사용 기한을 적어 본 메일로 회신 주시기 바랍니다.\n\n` +
      `감사합니다.`
    );
  } else {
    return (
      `${item.requester}${rank}님, 안녕하세요. 인사기획팀입니다.\n\n` +
      `이전에 신청하셨던 ${item.borrower}의 외주 명함 사용 기한[${fmtDate(item.due_date)}]이 지나 연락드립니다.\n\n` +
      `잔여 명함은 개인이 직접 파기 처리 부탁드립니다.\n\n` +
      `명함 미파기로 인한 정보 도용 등의 문제 발생 시 법적 책임이 따를 수 있으니 주의 바랍니다.\n\n` +
      `감사합니다.`
    );
  }
}

async function main() {
  console.log(`[${today}] 기한지연 자동 알림 시작`);

  const query = [
    'completed_date=is.null',
    `due_date=lt.${today}`,
    'requester_email=not.is.null',
    'requester_email=neq.',
    'last_notified=is.null',
    'select=*',
  ].join('&');

  const items = await supabaseGet(query);

  if (!items.length) {
    console.log('최초 발송할 항목 없음');
    return;
  }

  console.log(`발송 대상: ${items.length}건`);

  let success = 0, fail = 0;
  for (const item of items) {
    try {
      await sendEmail(item, buildMessage(item));
      await supabasePatch(item.id, { last_notified: today });
      console.log(`✓ ${item.requester} → ${item.requester_email}`);
      success++;
    } catch (e) {
      console.error(`✗ ${item.requester}: ${e.message}`);
      fail++;
    }
  }

  console.log(`\n완료: 성공 ${success}건 / 실패 ${fail}건`);
}

main().catch(e => {
  console.error('스크립트 오류:', e);
  process.exit(1);
});
