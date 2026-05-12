// scripts/auto-notify.js
// GitHub Actions에서 매일 자동 실행되는 기한지연 메일 발송 스크립트

import { createClient } from '@supabase/supabase-js';

const supabase = createClient(
  process.env.VITE_SUPABASE_URL,
  process.env.VITE_SUPABASE_ANON
);

const today = new Date().toISOString().slice(0, 10);

function fmtDate(d) {
  if (!d) return '-';
  const dt = new Date(d);
  return `${dt.getFullYear()}.${String(dt.getMonth()+1).padStart(2,'0')}.${String(dt.getDate()).padStart(2,'0')}`;
}

async function sendEmail(item) {
  const rank   = item.requester_rank ? ' ' + item.requester_rank : '';
  const message =
    `${item.requester}${rank}님, 안녕하세요. 인사기획팀입니다.\n\n` +
    `이전에 신청하셨던 ${item.borrower} 메일 계정[${item.item_name}]의 사용 기한[${fmtDate(item.due_date)}]이 지나 연락드립니다.\n\n` +
    `따라서 계정 비활성화 처리를 하고자 하는데 진행해도 되는지에 대하여 회신 부탁드립니다.\n\n` +
    `만약 연장이 필요할 경우, 연장 사유와 사용 기한을 적어 본 메일로 회신 주시기 바랍니다.\n\n` +
    `감사합니다.`;

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

  if (!res.ok) {
    const text = await res.text();
    throw new Error(text);
  }
}

async function main() {
  console.log(`[${today}] 기한지연 자동 알림 시작`);

  // 기한지연 항목 조회
  // 조건: 메일계정 + 완료 안됨 + 기한 지남 + 신청인 이메일 있음 + 한 번도 발송 안 함
  const { data: items, error } = await supabase
    .from('rentals')
    .select('*')
    .eq('category', '메일계정')
    .is('completed_date', null)
    .lt('due_date', today)
    .not('requester_email', 'is', null)
    .neq('requester_email', '')
    .is('last_notified', null);  // 최초 1회만

  if (error) {
    console.error('Supabase 조회 실패:', error.message);
    process.exit(1);
  }

  const targets = items;

  if (!targets.length) {
    console.log('최초 발송할 항목 없음 (이미 발송했거나 해당 없음)');
    return;
  }

  console.log(`발송 대상: ${targets.length}건`);

  let success = 0, fail = 0;

  for (const item of targets) {
    try {
      await sendEmail(item);
      // 발송 성공 시 last_notified 업데이트
      await supabase
        .from('rentals')
        .update({ last_notified: today })
        .eq('id', item.id);
      console.log(`✓ ${item.requester} → ${item.requester_email} (${item.item_name})`);
      success++;
    } catch (e) {
      console.error(`✗ ${item.requester} (${item.requester_email}): ${e.message}`);
      fail++;
    }
  }

  console.log(`\n완료: 성공 ${success}건 / 실패 ${fail}건`);
}

main().catch(e => {
  console.error('스크립트 오류:', e);
  process.exit(1);
});
