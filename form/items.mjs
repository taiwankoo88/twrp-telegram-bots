// Rules shared by the bot, the web server on the host PC and the static order page on GitHub Pages.
// Plain JavaScript only: this file also runs in the browser, so it must not import Node or Excel modules.

export const literal = s => /^[=+@-]/.test(s) ? "'" + s : s;

// Designation, standard, strength and remark for an item, shared by order files and weekly summaries.
export function describeItem(material, item) {
  const remark = item.parts.length > 1 ? `Combined: ${item.parts.join(' + ')}` : item.perSet || null;
  // Gasket wording follows the gasket template; the PN has its own column in the order file.
  if (item.kind === 'gasket') return {
    designation: `SIZE: DN${item.dn} ${item.thickness}MM THK Full-Face with Bolt Holes`,
    standard: 'BS 1154 Class Y3',
    grade: `PN${item.pn}`,
    remark,
  };
  const label = {bolt:'FULL THREADED HEXAGON BOLT', nut:'HEXAGON NUT',washer:'FLAT WASHER',rod:'STUD ROD'}[item.kind];
  const suffix = item.kind === 'washer' ? ` X ${item.length}mm thk.` : item.length ? `, L=${item.length}mm` : '';
  return {
    designation: `${label}: M${item.diameter}${suffix}`,
    standard: material === 'SS316' ? 'BS EN ISO 3506' : item.kind === 'washer' ? 'BS 4320' : 'BS EN ISO 4190',
    grade: material === 'HDG' ? 'G 8.8' : item.kind === 'nut' ? 'A4-80' : 'A4-70',
    remark,
  };
}

// Title and item lines of one section, as listed in weekly summaries, administrator notices and emails.
export function sectionLines(material, section) {
  return [section.title, ...section.items.map(item => {
    const {designation, grade} = describeItem(material, item);
    // A gasket's PN is its own column in the order file, so add it here.
    return `• ${designation}${item.kind === 'gasket' ? `, ${grade}` : ''} × ${item.qty.toLocaleString('en-US')}`;
  })];
}

// Subject: project, material, requester and titles, e.g. C3D1 HDG Fasteners · Nicholas Koo · WDL12 Chemical Hanging Piperack.
export function mailSubject(order) {
  const kind = order.material === 'EPDM' ? 'EPDM Gaskets' : `${order.material} Fasteners`;
  const subject = [`${order.project} ${kind}`, order.requester, order.sections.map(s => s.title).join('; ')].join(' · ');
  // Exchange rejects subjects longer than 255 characters.
  return subject.length <= 255 ? subject : `${subject.slice(0, 254)}…`;
}

const one = s => String(s ?? '').replace(/\s+/g, ' ').trim();
// The form sends numbers; anything else must not reach the text order (it could add lines of its own).
function num(value, what, where, whole = true) {
  const n = Number(value);
  if (value === '' || value === null || value === undefined || !(n > 0) || (whole && !Number.isInteger(n))) throw Error(`${what} is missing or not a ${whole ? 'whole ' : ''}number (${where}).`);
  return n;
}

// The web form as the bot's text order format, e.g. "C3D1 SS316 / Title: … / M16x110mm bolt = 2880",
// so that parseOrder applies every rule to web orders as well.
export function orderText(form) {
  if (!['C3D1', 'C3D2'].includes(form?.project)) throw Error('Choose the project: C3D1 or C3D2.');
  if (!['SS316', 'HDG', 'EPDM'].includes(form.material)) throw Error('Choose the material: SS316, HDG or EPDM gasket.');
  if (!one(form.requester)) throw Error('Enter your name as the requester.');
  const lines = [`${form.project} ${form.material === 'EPDM' ? 'Gasket' : form.material}`, `Requester Name: ${one(form.requester)}`];
  if (!Array.isArray(form.sections) || !form.sections.length) throw Error('Add at least one location / title.');
  for (const [i, sec] of form.sections.entries()) {
    const title = one(sec?.title);
    if (!title) throw Error(`Enter the location / title of section ${i + 1}.`);
    lines.push(`Title: ${title}`);
    if (sec.sets !== '' && sec.sets !== undefined && sec.sets !== null && Number(sec.sets) !== 1) lines.push(`Sets: ${num(sec.sets, 'Number of sets', title)}`);
    if (!Array.isArray(sec.items) || !sec.items.length) throw Error(`Add at least one item under "${title}".`);
    for (const [j, item] of sec.items.entries()) {
      const where = `${title}, row ${j + 1}`, qty = num(item?.qty, 'Quantity', where);
      if (form.material === 'EPDM') {
        lines.push(`DN${num(item.dn, 'DN', where)} PN${num(item.pn, 'PN', where, false)} ${num(item.thickness || 3, 'Thickness', where, false)}mm = ${qty}`);
        continue;
      }
      const d = num(item.diameter, 'Size', where);
      if (item.kind === 'nut') lines.push(`M${d} nut = ${qty}`);
      else if (item.kind === 'washer') lines.push(`M${d}${item.thickness ? `x${num(item.thickness, 'Washer thickness', where, false)}mm` : ''} washer = ${qty}`);
      else if (item.kind === 'bolt' || item.kind === 'rod') lines.push(`M${d}x${num(item.length, 'Length', where)}mm ${item.kind === 'rod' ? 'stud rod' : 'bolt'} = ${qty}`);
      else throw Error(`Choose bolt, nut, washer or stud rod (${where}).`);
    }
  }
  return lines.join('\n');
}
