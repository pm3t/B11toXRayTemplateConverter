import React, { useState } from 'react';
import {
  Coins,
  Check,
  Copy,
  ExternalLink,
  Gift,
  X,
  Sparkles,
  ArrowRight,
  MessageCircle,
  ShieldAlert,
  Lock,
} from 'lucide-react';
import { TOKEN_PACKAGES, BANK_ACCOUNTS, ADMIN_WHATSAPP_NUMBER } from '../config/pricing';
import { TokenPackage, UserWallet } from '../types/token';
import { redeemVoucherCode, getOrCreateWallet } from '../services/tokenService';

interface Props {
  isOpen: boolean;
  onClose: () => void;
  wallet?: UserWallet;
  onWalletUpdated?: () => void;
  onOpenAdminModal?: () => void;
  initialNotice?: string | null;
}

export default function TopUpModal({
  isOpen,
  onClose,
  wallet,
  onWalletUpdated,
  onOpenAdminModal,
  initialNotice,
}: Props) {
  const activeWallet = wallet || getOrCreateWallet();
  const [selectedPkg, setSelectedPkg] = useState<TokenPackage>(TOKEN_PACKAGES[1]); // default 50 tokens
  const [selectedBank, setSelectedBank] = useState<string>(BANK_ACCOUNTS[0].bank);
  const [copiedBank, setCopiedBank] = useState<string | null>(null);
  const [copiedUserId, setCopiedUserId] = useState(false);
  const [voucherCode, setVoucherCode] = useState('');
  const [redeemStatus, setRedeemStatus] = useState<{ success: boolean; message: string } | null>(null);
  const [activeTab, setActiveTab] = useState<'buy' | 'redeem'>('buy');

  if (!isOpen) return null;

  const copyToClipboard = (text: string, type: 'bank' | 'userId', bankName?: string) => {
    navigator.clipboard.writeText(text);
    if (type === 'bank' && bankName) {
      setCopiedBank(bankName);
      setTimeout(() => setCopiedBank(null), 2000);
    } else if (type === 'userId') {
      setCopiedUserId(true);
      setTimeout(() => setCopiedUserId(false), 2000);
    }
  };

  const handleRedeem = (e: React.FormEvent) => {
    e.preventDefault();
    if (!voucherCode.trim()) return;

    const res = redeemVoucherCode(voucherCode);
    if (res.success) {
      setRedeemStatus({
        success: true,
        message: `Selamat! Voucher berhasil ditukarkan. Saldo Anda bertambah +${res.tokensAdded} Token (Total: ${res.newBalance} Token).`,
      });
      setVoucherCode('');
      onWalletUpdated?.();
    } else {
      setRedeemStatus({
        success: false,
        message: res.error || 'Kode voucher tidak valid atau sudah kedaluwarsa.',
      });
    }
  };

  const currentBank = BANK_ACCOUNTS.find((b) => b.bank === selectedBank) || BANK_ACCOUNTS[0];

  const waMessage = encodeURIComponent(
    `Halo, saya ingin konfirmasi pembelian token X-Ray Generator:\n\n` +
      `• Paket: ${selectedPkg.tokens} Token (${selectedPkg.formattedPrice})\n` +
      `• User ID: ${activeWallet.userId}\n` +
      `• Bank Tujuan: ${currentBank.bank} (${currentBank.number} a.n ${currentBank.holder})\n\n` +
      `Berikut saya lampirkan bukti transfer. Mohon dikirimkan kode voucher tokennya. Terima kasih!`
  );

  const waUrl = `https://wa.me/${ADMIN_WHATSAPP_NUMBER}?text=${waMessage}`;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/75 backdrop-blur-md animate-in fade-in duration-200">
      <div className="relative w-full max-w-2xl bg-slate-900 border border-slate-700/80 rounded-3xl shadow-2xl overflow-hidden flex flex-col max-h-[90vh]">
        {/* Header Modal */}
        <div className="p-6 bg-gradient-to-r from-slate-900 via-indigo-950/40 to-slate-900 border-b border-slate-800 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-11 h-11 rounded-2xl bg-gradient-to-tr from-amber-500 to-amber-300 text-slate-950 flex items-center justify-center shadow-lg shadow-amber-500/20 font-bold">
              <Coins className="w-6 h-6" />
            </div>
            <div>
              <h2 className="text-lg font-bold text-white flex items-center gap-2">
                Saldo &amp; Top Up Token
              </h2>
              <div className="flex items-center gap-2 text-xs text-slate-400">
                <span>
                  Sisa Saldo: <b className="text-amber-400 font-semibold">{activeWallet.balance} Token</b>
                </span>
                <span>&bull;</span>
                <span className="flex items-center gap-1">
                  ID: <code className="text-slate-300 font-mono">{activeWallet.userId}</code>
                  <button
                    onClick={() => copyToClipboard(activeWallet.userId, 'userId')}
                    className="hover:text-white p-0.5"
                    title="Salin ID"
                  >
                    {copiedUserId ? <Check className="w-3 h-3 text-emerald-400" /> : <Copy className="w-3 h-3" />}
                  </button>
                </span>
              </div>
            </div>
          </div>
          <button
            onClick={onClose}
            className="text-slate-400 hover:text-white p-2 rounded-full hover:bg-slate-800 transition"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Notice alert jika dipicu karena token habis */}
        {initialNotice && (
          <div className="px-6 py-3 bg-amber-500/10 border-b border-amber-500/20 flex items-center gap-2.5 text-amber-300 text-xs font-medium">
            <ShieldAlert className="w-4 h-4 shrink-0 text-amber-400" />
            <span>{initialNotice}</span>
          </div>
        )}

        {/* Tab Switcher */}
        <div className="px-6 pt-4 flex border-b border-slate-800 gap-4 text-sm font-medium">
          <button
            onClick={() => setActiveTab('buy')}
            className={`pb-3 px-1 border-b-2 transition flex items-center gap-2 ${
              activeTab === 'buy'
                ? 'border-indigo-500 text-indigo-400 font-semibold'
                : 'border-transparent text-slate-400 hover:text-slate-200'
            }`}
          >
            <Coins className="w-4 h-4" /> Beli Token Baru
          </button>
          <button
            onClick={() => setActiveTab('redeem')}
            className={`pb-3 px-1 border-b-2 transition flex items-center gap-2 ${
              activeTab === 'redeem'
                ? 'border-indigo-500 text-indigo-400 font-semibold'
                : 'border-transparent text-slate-400 hover:text-slate-200'
            }`}
          >
            <Gift className="w-4 h-4" /> Tukar Kode Voucher
          </button>
        </div>

        {/* Content Body */}
        <div className="p-6 overflow-y-auto space-y-6 flex-1 text-slate-200 text-sm">
          {activeTab === 'buy' ? (
            <>
              {/* 1. Pilih Paket */}
              <div>
                <label className="block text-xs font-semibold text-slate-400 uppercase tracking-wider mb-2.5">
                  1. Pilih Paket Token
                </label>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                  {TOKEN_PACKAGES.map((pkg) => (
                    <div
                      key={pkg.id}
                      onClick={() => setSelectedPkg(pkg)}
                      className={`relative cursor-pointer p-4 rounded-2xl border transition-all ${
                        selectedPkg.id === pkg.id
                          ? 'bg-indigo-600/15 border-indigo-500 shadow-lg shadow-indigo-500/10 scale-[1.02]'
                          : 'bg-slate-800/40 border-slate-700/80 hover:border-slate-600'
                      }`}
                    >
                      {pkg.isPopular && (
                        <span className="absolute -top-2.5 right-3 bg-gradient-to-r from-amber-500 to-orange-500 text-slate-950 font-bold text-[10px] uppercase px-2 py-0.5 rounded-full shadow">
                          Terpopuler
                        </span>
                      )}
                      <div className="font-bold text-lg text-white mb-0.5">{pkg.tokens} Token</div>
                      <div className="text-indigo-400 font-semibold text-base mb-1.5">
                        {pkg.formattedPrice}
                      </div>
                      <p className="text-[11px] text-slate-400 leading-relaxed">{pkg.description}</p>
                    </div>
                  ))}
                </div>
              </div>

              {/* 2. Pilihan Bank Transfer */}
              <div>
                <label className="block text-xs font-semibold text-slate-400 uppercase tracking-wider mb-2.5">
                  2. Transfer ke Rekening Bank
                </label>
                <div className="grid grid-cols-3 gap-2 mb-3">
                  {BANK_ACCOUNTS.map((acc) => (
                    <button
                      key={acc.bank}
                      type="button"
                      onClick={() => setSelectedBank(acc.bank)}
                      className={`py-2 px-3 rounded-xl border font-bold text-xs flex items-center justify-center gap-1.5 transition ${
                        selectedBank === acc.bank
                          ? 'bg-slate-800 border-indigo-500 text-indigo-400 shadow'
                          : 'bg-slate-800/40 border-slate-700 text-slate-400 hover:text-slate-200'
                      }`}
                    >
                      <span>Bank {acc.bank}</span>
                    </button>
                  ))}
                </div>

                {/* Detail Rekening Terpilih */}
                <div className="p-4 rounded-2xl bg-slate-950/80 border border-slate-800 space-y-3">
                  <div className="flex items-center justify-between">
                    <div>
                      <span className="text-[11px] text-slate-400 block font-medium">Bank Tujuan</span>
                      <span className="font-bold text-white text-sm">Bank {currentBank.bank}</span>
                    </div>
                    <div className="text-right">
                      <span className="text-[11px] text-slate-400 block font-medium">Atas Nama</span>
                      <span className="font-semibold text-slate-200 text-sm">{currentBank.holder}</span>
                    </div>
                  </div>

                  <div className="p-3 rounded-xl bg-slate-900 border border-slate-800 flex items-center justify-between">
                    <div>
                      <span className="text-[10px] text-slate-400 uppercase block font-semibold">
                        Nomor Rekening
                      </span>
                      <span className="font-mono text-base font-bold text-amber-400 tracking-wider">
                        {currentBank.number}
                      </span>
                    </div>
                    <button
                      type="button"
                      onClick={() => copyToClipboard(currentBank.number, 'bank', currentBank.bank)}
                      className="px-3 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white font-medium text-xs flex items-center gap-1.5 transition shadow"
                    >
                      {copiedBank === currentBank.bank ? (
                        <>
                          <Check className="w-3.5 h-3.5" /> Tersalin!
                        </>
                      ) : (
                        <>
                          <Copy className="w-3.5 h-3.5" /> Salin Rekening
                        </>
                      )}
                    </button>
                  </div>

                  <div className="flex items-center justify-between text-xs pt-1 border-t border-slate-900">
                    <span className="text-slate-400">Total Pembayaran:</span>
                    <span className="font-bold text-emerald-400 text-sm">
                      {selectedPkg.formattedPrice}
                    </span>
                  </div>
                </div>
              </div>

              {/* 3. Tombol WhatsApp Konfirmasi */}
              <div>
                <label className="block text-xs font-semibold text-slate-400 uppercase tracking-wider mb-2">
                  3. Konfirmasi &amp; Terima Kode Voucher
                </label>
                <a
                  href={waUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="w-full py-3.5 px-4 rounded-2xl bg-emerald-600 hover:bg-emerald-500 text-white font-bold flex items-center justify-center gap-2 shadow-lg shadow-emerald-600/25 transition-all text-sm group"
                >
                  <MessageCircle className="w-5 h-5 fill-current" />
                  Konfirmasi Pembayaran via WhatsApp
                  <ExternalLink className="w-4 h-4 ml-1 opacity-70 group-hover:opacity-100" />
                </a>
                <p className="text-[11px] text-slate-400 mt-2 text-center">
                  Setelah mengirim bukti transfer via WhatsApp, Anda akan segera menerima{' '}
                  <span className="text-slate-200 font-semibold">Kode Voucher</span> untuk diisi pada tab "Tukar Kode
                  Voucher".
                </p>
              </div>
            </>
          ) : (
            /* TAB REDEEM VOUCHER */
            <div className="space-y-5 py-2">
              <div className="p-4 rounded-2xl bg-indigo-500/10 border border-indigo-500/20 text-indigo-300 text-xs leading-relaxed flex items-start gap-3">
                <Sparkles className="w-5 h-5 text-indigo-400 shrink-0 mt-0.5" />
                <div>
                  <span className="font-bold block text-sm text-indigo-200 mb-0.5">Punya Kode Voucher?</span>
                  Masukkan kode voucher token yang Anda terima dari Admin setelah melakukan transfer pembayaran. Saldo
                  akan langsung bertambah otomatis.
                </div>
              </div>

              <form onSubmit={handleRedeem} className="space-y-4">
                <div>
                  <label className="block text-xs font-semibold text-slate-300 mb-1.5">
                    Masukkan Kode Voucher Token
                  </label>
                  <input
                    type="text"
                    value={voucherCode}
                    onChange={(e) => setVoucherCode(e.target.value.toUpperCase())}
                    placeholder="Contoh: XRAY-50-A7B2-9F12"
                    className="w-full px-4 py-3 bg-slate-950 border border-slate-700 rounded-xl text-base font-mono uppercase tracking-wider placeholder:normal-case placeholder:font-sans placeholder:text-slate-500 focus:outline-none focus:border-indigo-500"
                  />
                </div>

                {redeemStatus && (
                  <div
                    className={`p-3.5 rounded-xl text-xs font-medium ${
                      redeemStatus.success
                        ? 'bg-emerald-500/15 border border-emerald-500/30 text-emerald-300'
                        : 'bg-rose-500/15 border border-rose-500/30 text-rose-300'
                    }`}
                  >
                    {redeemStatus.message}
                  </div>
                )}

                <button
                  type="submit"
                  disabled={!voucherCode.trim()}
                  className="w-full py-3 rounded-xl bg-gradient-to-r from-indigo-500 to-cyan-500 hover:from-indigo-400 hover:to-cyan-400 disabled:opacity-50 text-white font-bold text-sm shadow-lg shadow-indigo-500/25 transition flex items-center justify-center gap-2"
                >
                  <Gift className="w-4 h-4" /> Tukarkan Voucher Sekarang
                </button>
              </form>
            </div>
          )}
        </div>

        {/* Footer Admin Shortcut */}
        <div className="px-6 py-3 bg-slate-950 border-t border-slate-800/80 flex items-center justify-between text-xs text-slate-500">
          <span>1x Konversi = 1 Token</span>
          <button
            onClick={() => {
              onClose();
              onOpenAdminModal?.();
            }}
            className="text-slate-400 hover:text-amber-400 flex items-center gap-1 font-medium transition"
          >
            <Lock className="w-3.5 h-3.5" /> Generator Voucher Admin
          </button>
        </div>
      </div>
    </div>
  );
}
