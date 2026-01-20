/**
 * @fileoverview Gerenciador de Cache para Google Apps Script
 * @version 1.1.0
 *
 * Fornece uma interface simplificada para CacheService com suporte
 * a objetos JSON e chunking para dados maiores que 100KB.
 *
 * NOTA: Renomeado para SharedCache para evitar conflito com
 * CacheManager do BOM.gs que tem logica especifica
 *
 * USO:
 *   SharedCache.set('key', { dados: 'valor' }, 300);
 *   const data = SharedCache.get('key');
 *   const data = SharedCache.getOrFetch('key', () => fetchData(), 300);
 */

/**
 * Gerenciador de cache compartilhado
 * @namespace
 */
const SharedCache = {
  /** @private */
  _cache: CacheService.getScriptCache(),

  /** TTL padrão em segundos (5 minutos) */
  DEFAULT_TTL: 300,

  /** Tamanho máximo de chunk (90KB para margem de segurança) */
  MAX_CHUNK_SIZE: 90000,

  /**
   * Obtém valor do cache
   * @param {string} key - Chave do cache
   * @returns {*} Valor parseado ou null se não existir
   */
  get(key) {
    const data = this._cache.get(key);
    if (!data) return null;

    try {
      return JSON.parse(data);
    } catch (e) {
      return data;
    }
  },

  /**
   * Salva valor no cache
   * @param {string} key - Chave do cache
   * @param {*} value - Valor a salvar (será convertido para JSON)
   * @param {number} [ttl] - Tempo de vida em segundos (max 21600)
   */
  set(key, value, ttl = this.DEFAULT_TTL) {
    const data = typeof value === 'string' ? value : JSON.stringify(value);
    this._cache.put(key, data, Math.min(ttl, 21600));
  },

  /**
   * Remove valor do cache
   * @param {string} key - Chave a remover
   */
  remove(key) {
    this._cache.remove(key);
  },

  /**
   * Remove múltiplas chaves
   * @param {string[]} keys - Array de chaves
   */
  removeAll(keys) {
    this._cache.removeAll(keys);
  },

  /**
   * Obtém dados com fallback para função de fetch
   * @param {string} key - Chave do cache
   * @param {Function} fetchFn - Função para buscar dados se cache miss
   * @param {number} [ttl] - TTL para novos dados
   * @returns {*} Dados do cache ou da função
   */
  getOrFetch(key, fetchFn, ttl = this.DEFAULT_TTL) {
    let data = this.get(key);

    if (data === null) {
      data = fetchFn();
      this.set(key, data, ttl);
    }

    return data;
  },

  /**
   * Salva dados grandes em chunks (para dados > 100KB)
   * @param {string} key - Chave base
   * @param {*} value - Valor a salvar
   * @param {number} [ttl] - TTL em segundos
   */
  setLarge(key, value, ttl = this.DEFAULT_TTL) {
    const data = JSON.stringify(value);
    const chunks = this._chunkString(data, this.MAX_CHUNK_SIZE);

    chunks.forEach((chunk, i) => {
      this._cache.put(`${key}_chunk_${i}`, chunk, ttl);
    });

    this._cache.put(`${key}_meta`, JSON.stringify({
      chunks: chunks.length,
      timestamp: Date.now()
    }), ttl);
  },

  /**
   * Recupera dados grandes de chunks
   * @param {string} key - Chave base
   * @returns {*} Dados reconstruídos ou null
   */
  getLarge(key) {
    const metaStr = this._cache.get(`${key}_meta`);
    if (!metaStr) return null;

    try {
      const meta = JSON.parse(metaStr);
      const chunks = [];

      for (let i = 0; i < meta.chunks; i++) {
        const chunk = this._cache.get(`${key}_chunk_${i}`);
        if (!chunk) return null;
        chunks.push(chunk);
      }

      return JSON.parse(chunks.join(''));
    } catch (e) {
      console.error('Erro ao recuperar dados grandes:', e);
      return null;
    }
  },

  /**
   * Remove dados grandes (chunks)
   * @param {string} key - Chave base
   */
  removeLarge(key) {
    const metaStr = this._cache.get(`${key}_meta`);
    if (!metaStr) return;

    try {
      const meta = JSON.parse(metaStr);
      const keys = [`${key}_meta`];

      for (let i = 0; i < meta.chunks; i++) {
        keys.push(`${key}_chunk_${i}`);
      }

      this._cache.removeAll(keys);
    } catch (e) {
      console.error('Erro ao remover dados grandes:', e);
    }
  },

  /**
   * Divide string em chunks
   * @private
   */
  _chunkString(str, size) {
    const chunks = [];
    for (let i = 0; i < str.length; i += size) {
      chunks.push(str.slice(i, i + size));
    }
    return chunks;
  }
};

// Exportar para uso global
if (typeof module !== 'undefined') {
  module.exports = SharedCache;
}
