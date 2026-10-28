const equal = (a, b) => JSON.stringify(a) === JSON.stringify(b);
const object = value => value !== null && typeof value === 'object' && !Array.isArray(value);
const forbidden = new Set(['__proto__', 'prototype', 'constructor']);

// Three-way merge: unrelated changes from two computers survive; same-field edits conflict.
function mergeState(base, local, remote, path = '') {
  if (equal(local, base)) return structuredClone(remote);
  if (equal(remote, base) || equal(local, remote)) return structuredClone(local);
  if (object(base) && object(local) && object(remote)) {
    const result = {};
    for (const key of new Set([...Object.keys(base), ...Object.keys(local), ...Object.keys(remote)])) {
      if (forbidden.has(key)) throw new Error('Invalid field');
      const merged = mergeState(base[key], local[key], remote[key], path + '/' + key);
      if (merged !== undefined) result[key] = merged;
    }
    return result;
  }
  throw new Error('동시에 변경된 항목: ' + path);
}

class PostgresStateStore {
  constructor(pool) { this.pool = pool; }
  async initialize() {
    await this.pool.query(`CREATE TABLE IF NOT EXISTS payroll_shared_state (
      id integer PRIMARY KEY CHECK (id = 1), revision bigint NOT NULL,
      payload jsonb NOT NULL, updated_at timestamptz NOT NULL DEFAULT now())`);
  }
  async read() {
    const { rows } = await this.pool.query('SELECT revision, payload, updated_at FROM payroll_shared_state WHERE id=1');
    return rows.length ? {revision:Number(rows[0].revision), payload:rows[0].payload, updatedAt:rows[0].updated_at} : {revision:0,payload:null};
  }
  async write(revision, payload) {
    const query = revision === 0
      ? `INSERT INTO payroll_shared_state (id,revision,payload) VALUES (1,$1::bigint+1,$2::jsonb)
         ON CONFLICT (id) DO NOTHING RETURNING revision,updated_at`
      : `UPDATE payroll_shared_state SET revision=revision+1,payload=$2::jsonb,updated_at=now()
         WHERE id=1 AND revision=$1::bigint RETURNING revision,updated_at`;
    const { rows } = await this.pool.query(query, [revision,JSON.stringify(payload)]);
    if (rows.length) return {revision:Number(rows[0].revision),updatedAt:rows[0].updated_at};
    return null;
  }
}
module.exports = { mergeState, PostgresStateStore };
