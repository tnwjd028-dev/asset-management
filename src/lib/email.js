import emailjs from '@emailjs/browser'

const PUBLIC_KEY  = import.meta.env.VITE_EMAILJS_PUBLIC_KEY
const SERVICE_ID  = import.meta.env.VITE_EMAILJS_SERVICE_ID
const TEMPLATE_ID = import.meta.env.VITE_EMAILJS_TEMPLATE_ID

emailjs.init(PUBLIC_KEY)

export async function sendNotification(item, message) {
  const toEmail = item.requesterEmail
  if (!toEmail) throw new Error('신청인 이메일이 없습니다.')
  return emailjs.send(SERVICE_ID, TEMPLATE_ID, {
    to_email:  toEmail,
    to_name:   item.requester,
    category:  item.category,
    item_name: item.itemName,
    due_date:  item.dueDate,
    purpose:   item.purpose || '-',
    message:   message,
    from_name: 'Bimatrix 자산관리팀',
  })
}

export async function sendBulkNotifications(items, buildMessage) {
  let success = 0, fail = 0
  const errors = []
  for (const item of items) {
    try {
      await sendNotification(item, buildMessage(item))
      success++
    } catch (e) {
      fail++
      const msg = e?.text || e?.message || JSON.stringify(e)
      errors.push(`${item.requester}(${item.requesterEmail || '이메일 없음'}): ${msg}`)
    }
  }
  return { success, fail, errors }
}
