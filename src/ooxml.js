// Namespace-aware element and attribute access for WordprocessingML, Transitional and Strict.
export const W = new Set(['http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'http://purl.oclc.org/ooxml/wordprocessingml/main']);
export const isW = (node, name) => node?.nodeType === 1 && W.has(node.namespaceURI) && (!name || node.localName === name);
export const kids = node => [...(node?.children || [])];
export const child = (node, name) => kids(node).find(el => isW(el, name));
export const attr = (node, name) => {
  if (!node) return '';
  for (const ns of W) if (node.hasAttributeNS(ns, name)) return node.getAttributeNS(ns, name);
  return '';
};
export const descendants = (node, name) => [...(node?.getElementsByTagNameNS('*', name) || [])].filter(el => W.has(el.namespaceURI));
export const on = node => !!node && !['0', 'false', 'off'].includes(attr(node, 'val'));
