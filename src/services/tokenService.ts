import { UserWallet, TokenTransaction } from '../types/token';
import { FREE_WELCOME_TOKENS, CONVERSION_TOKEN_COST } from '../config/pricing';

const STORAGE_KEY = 'xray_user_wallet_v1';
const REDEEMED_CODES_KEY = 'xray_redeemed_codes_v1';
const VOUCHER_SALT = 'XRAY_SECURE_TOKEN_SALT_2026';

function generateRandomId(prefix = 'USR'): string {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let result = '';
  for (let i = 0; i < 6; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return `${prefix}-${result}`;
}

// Checksum sederhana untuk validasi voucher tanpa perlu server
function calculateChecksum(payload: string): string {
  let hash = 2166136261;
  const combined = payload + VOUCHER_SALT;
  for (let i = 0; i < combined.length; i++) {
    hash ^= combined.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return Math.abs(hash % 65536).toString(16).toUpperCase().padStart(4, '0');
}

export function generateVoucher(tokens: number): string {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let rand = '';
  for (let i = 0; i < 4; i++) {
    rand += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  const payload = `${tokens}-${rand}`;
  const checksum = calculateChecksum(payload);
  return `XRAY-${tokens}-${rand}-${checksum}`;
}

export function validateVoucher(code: string): { valid: boolean; tokens: number; error?: string } {
  const cleanCode = code.trim().toUpperCase();
  const parts = cleanCode.split('-');
  
  if (parts.length !== 4 || parts[0] !== 'XRAY') {
    return { valid: false, tokens: 0, error: 'Format kode voucher tidak valid (Contoh: XRAY-50-ABCD-1234)' };
  }

  const tokens = parseInt(parts[1], 10);
  const rand = parts[2];
  const checksum = parts[3];

  if (isNaN(tokens) || tokens <= 0) {
    return { valid: false, tokens: 0, error: 'Nominal token pada voucher tidak valid.' };
  }

  const expectedChecksum = calculateChecksum(`${tokens}-${rand}`);
  if (checksum !== expectedChecksum) {
    return { valid: false, tokens: 0, error: 'Kode voucher salah atau tidak terverifikasi.' };
  }

  // Cek apakah sudah pernah digunakan pada browser ini
  const redeemed = getRedeemedCodes();
  if (redeemed.includes(cleanCode)) {
    return { valid: false, tokens: 0, error: 'Kode voucher ini sudah pernah digunakan.' };
  }

  return { valid: true, tokens };
}

function getRedeemedCodes(): string[] {
  try {
    const raw = localStorage.getItem(REDEEMED_CODES_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch {
    return [];
  }
}

function markCodeRedeemed(code: string) {
  try {
    const redeemed = getRedeemedCodes();
    redeemed.push(code.trim().toUpperCase());
    localStorage.setItem(REDEEMED_CODES_KEY, JSON.stringify(redeemed));
  } catch (e) {
    console.error('Gagal mencatat voucher:', e);
  }
}

export function getOrCreateWallet(): UserWallet {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed: UserWallet = JSON.parse(raw);
      // Validasi data dasar
      if (typeof parsed.balance === 'number' && parsed.userId) {
        return parsed;
      }
    }
  } catch (e) {
    console.warn('Gagal membaca wallet dari storage, membuat wallet baru:', e);
  }

  // Inisialisasi wallet baru untuk user yang baru join dengan 10 FREE TOKEN
  const newWallet: UserWallet = {
    userId: generateRandomId('USR'),
    balance: FREE_WELCOME_TOKENS,
    createdAt: new Date().toISOString(),
    totalConversions: 0,
    hasReceivedFreeBonus: true,
    transactions: [
      {
        id: 'tx-welcome',
        timestamp: new Date().toISOString(),
        type: 'bonus',
        amount: FREE_WELCOME_TOKENS,
        balanceAfter: FREE_WELCOME_TOKENS,
        description: `Bonus Selamat Datang (${FREE_WELCOME_TOKENS} Token Gratis)`,
      },
    ],
  };

  saveWallet(newWallet);
  return newWallet;
}

export function saveWallet(wallet: UserWallet): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(wallet));
    window.dispatchEvent(new CustomEvent('xray:wallet_updated', { detail: wallet }));
  } catch (e) {
    console.error('Gagal menyimpan wallet:', e);
  }
}

export function consumeToken(): { success: boolean; balance: number; error?: string } {
  const wallet = getOrCreateWallet();
  if (wallet.balance < CONVERSION_TOKEN_COST) {
    return {
      success: false,
      balance: wallet.balance,
      error: `Token Anda tidak mencukupi (Sisa: ${wallet.balance}). Diperlukan ${CONVERSION_TOKEN_COST} token untuk 1x konversi.`,
    };
  }

  const newBalance = wallet.balance - CONVERSION_TOKEN_COST;
  const tx: TokenTransaction = {
    id: `tx-conv-${Date.now()}`,
    timestamp: new Date().toISOString(),
    type: 'conversion',
    amount: -CONVERSION_TOKEN_COST,
    balanceAfter: newBalance,
    description: 'Konversi Dokumen Manifes (1 Token)',
  };

  wallet.balance = newBalance;
  wallet.totalConversions += 1;
  wallet.transactions.unshift(tx);
  saveWallet(wallet);

  return { success: true, balance: newBalance };
}

export function redeemVoucherCode(code: string): { success: boolean; tokensAdded: number; newBalance: number; error?: string } {
  const validation = validateVoucher(code);
  if (!validation.valid) {
    return { success: false, tokensAdded: 0, newBalance: 0, error: validation.error };
  }

  const wallet = getOrCreateWallet();
  const cleanCode = code.trim().toUpperCase();
  const newBalance = wallet.balance + validation.tokens;

  const tx: TokenTransaction = {
    id: `tx-redeem-${Date.now()}`,
    timestamp: new Date().toISOString(),
    type: 'redeem',
    amount: validation.tokens,
    balanceAfter: newBalance,
    description: `Top Up Voucher (+${validation.tokens} Token)`,
    referenceCode: cleanCode,
  };

  wallet.balance = newBalance;
  wallet.transactions.unshift(tx);
  saveWallet(wallet);
  markCodeRedeemed(cleanCode);

  return { success: true, tokensAdded: validation.tokens, newBalance };
}

export function adminAddTokens(amount: number, reason = 'Penyesuaian Admin'): { success: boolean; newBalance: number } {
  const wallet = getOrCreateWallet();
  const newBalance = Math.max(0, wallet.balance + amount);

  const tx: TokenTransaction = {
    id: `tx-admin-${Date.now()}`,
    timestamp: new Date().toISOString(),
    type: 'admin_adjustment',
    amount,
    balanceAfter: newBalance,
    description: `${reason} (${amount >= 0 ? '+' : ''}${amount} Token)`,
  };

  wallet.balance = newBalance;
  wallet.transactions.unshift(tx);
  saveWallet(wallet);

  return { success: true, newBalance };
}
