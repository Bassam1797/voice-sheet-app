export const storage = {
  get<T = any>(key: string): T | null {
    try {
      return JSON.parse(localStorage.getItem(key) || '');
    } catch {
      return null;
    }
  },
  set(key: string, val: any): void {
    localStorage.setItem(key, JSON.stringify(val));
  },
  remove(key: string): void {
    localStorage.removeItem(key);
  }
};
