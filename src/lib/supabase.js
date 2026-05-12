import { createClient } from '@supabase/supabase-js'

const SUPABASE_URL  = import.meta.env.VITE_SUPABASE_URL
const SUPABASE_ANON = import.meta.env.VITE_SUPABASE_ANON

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON)

export async function fetchAll() {
  const { data, error } = await supabase
    .from('rentals')
    .select('*')
    .order('created_at', { ascending: false })
  if (error) throw error
  return data.map(dbRowToItem)
}

export async function insertItem(item) {
  const { data, error } = await supabase
    .from('rentals')
    .insert(itemToDbRow(item))
    .select()
  if (error) throw error
  return dbRowToItem(data[0])
}

export async function updateItem(item) {
  const { data, error } = await supabase
    .from('rentals')
    .update(itemToDbRow(item))
    .eq('id', item.id)
    .select()
  if (error) throw error
  return dbRowToItem(data[0])
}

export async function deleteItem(id) {
  const { error } = await supabase.from('rentals').delete().eq('id', id)
  if (error) throw error
}

function itemToDbRow(item) {
  return {
    id:              item.id,
    category:        item.category,
    card_type:       item.cardType        || '',
    item_name:       item.itemName        || '',
    email_prefix:    item.emailPrefix     || '',
    requester:       item.requester       || '',
    requester_rank:  item.requesterRank   || '',
    requester_dept:  item.requesterDept   || '',
    requester_email: item.requesterEmail  || '',
    borrower:        item.borrower        || '',
    borrower_rank:   item.borrowerRank    || '',
    department:      item.department      || '',
    contact:         item.contact         || '',
    borrower_email:  item.borrowerEmail   || '',
    purpose:         item.purpose         || '',
    loan_date:       item.loanDate        || null,
    due_date:        item.dueDate         || null,
    notes:           item.notes           || '',
    completed_date:  item.completedDate   || null,
    security_file:   item.securityFile    || null,
    extensions:      item.extensions      || [],
  }
}

function dbRowToItem(row) {
  return {
    id:             row.id,
    category:       row.category,
    cardType:       row.card_type        || '',
    itemName:       row.item_name        || '',
    emailPrefix:    row.email_prefix     || '',
    requester:      row.requester        || '',
    requesterRank:  row.requester_rank   || '',
    requesterDept:  row.requester_dept   || '',
    requesterEmail: row.requester_email  || '',
    borrower:       row.borrower         || '',
    borrowerRank:   row.borrower_rank    || '',
    department:     row.department       || '',
    contact:        row.contact          || '',
    borrowerEmail:  row.borrower_email   || '',
    purpose:        row.purpose          || '',
    loanDate:       row.loan_date        || '',
    dueDate:        row.due_date         || '',
    notes:          row.notes            || '',
    completedDate:  row.completed_date   || null,
    securityFile:   row.security_file    || null,
    securityFileId: row.security_file ? 'db' : null,
    extensions:     row.extensions       || [],
  }
}
