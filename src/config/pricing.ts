import { BankAccount, TokenPackage } from '../types/token';

export const FREE_WELCOME_TOKENS = 10;
export const CONVERSION_TOKEN_COST = 1;

export const TOKEN_PACKAGES: TokenPackage[] = [
  {
    id: 'pkg-20',
    tokens: 20,
    price: 20000,
    formattedPrice: 'Rp 20.000',
    description: 'Rp 1.000 per konversi. Cocok untuk kebutuhan santai & berkala.',
  },
  {
    id: 'pkg-50',
    tokens: 50,
    price: 50000,
    formattedPrice: 'Rp 50.000',
    isPopular: true,
    description: 'Pilihan paling populer. Cukup untuk operasional reguler.',
  },
  {
    id: 'pkg-100',
    tokens: 100,
    price: 100000,
    formattedPrice: 'Rp 100.000',
    description: 'Nilai terbaik untuk tim cargo & logistik bervolume tinggi.',
  },
];

export const BANK_ACCOUNTS: BankAccount[] = [
  {
    bank: 'BCA',
    number: '6090260021',
    holder: 'Bayu Nurcahyono',
    logoColor: 'from-blue-600 to-blue-800',
  },
  {
    bank: 'Jago',
    number: '105692105503',
    holder: 'Bayu Nurcahyono',
    logoColor: 'from-amber-500 to-orange-600',
  },
  {
    bank: 'BRI',
    number: '140001000726504',
    holder: 'Bayu Nurcahyono',
    logoColor: 'from-blue-700 to-cyan-600',
  },
];

// Nomor WhatsApp admin untuk konfirmasi pembayaran
// Jika nomor belum diset, bisa disesuaikan atau diarahkan ke pesan WA umum
export const ADMIN_WHATSAPP_NUMBER = '6281234567890'; // Silakan sesuaikan dengan no WA Bayu

// PIN Rahasia Admin (bisa di-override lewat environment variable VITE_ADMIN_PIN)
export const ADMIN_SECRET_PIN = (import.meta.env.VITE_ADMIN_PIN as string) || '2026';
