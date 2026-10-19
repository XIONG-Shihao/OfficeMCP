export function readWindow(items, { query, offset, limit }) {
  const needle = query?.toLowerCase();
  const matching = needle
    ? items.filter((item) => JSON.stringify(item).toLowerCase().includes(needle))
    : items;
  return {
    items: matching.slice(offset, offset + limit),
    total: matching.length,
    offset,
    nextOffset: offset + limit < matching.length ? offset + limit : null,
  };
}
