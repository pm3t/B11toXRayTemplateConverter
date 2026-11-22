import React, { useState } from 'react';
import { Shield, Key, Copy, Check, Sparkles, X, PlusCircle } from 'lucide-react';
import { generateVoucher, adminAddTokens } from '../services/tokenService';
import { ADMIN_SECRET_PIN } from '../config/pricing';

interface Props {
  isOpen: boolean;
  onClose: () => void;
  onTokensUpdated: () => void;
}

export default function AdminVoucherModal({ isOpen, onClose, onTokensUpdated }: Props) {
  const [pin, setPin] = useState('');
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [pinError, setPinError] = useState(false);
  const [selectedTokens, setSelectedTokens] = useState<number>(50);
  const [generatedCode, setGeneratedCode] = useState<string>('');
  const [copied, setCopied] = useState(false);
  const [directAddAmount, setDirectAddAmount] = useState<number>(10);
  const [directAddSuccess, setDirectAddSuccess] = useState(false);

  if (!isOpen) return null;

  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    if (pin === ADMIN_SECRET_PIN) {
      setIsAuthenticated(true);
      setPinError(false);
    } else {
      setPinError(true);
    }
  };

  const handleGenerate = (tokens: number) => {
    const code = generateVoucher(tokens);
    setGeneratedCode(code);
    setCopied(false);
  };

  const copyToClipboard = (text: string) => {
    navigator.clipboard.writeText(text);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleDirectAdd = () => {
    adminAddTokens(directAddAmount, 'Manual Admin Credit');
    setDirectAddSuccess(true);
    onTokensUpdated();
    setTimeout(() => setDirectAddSuccess(false), 2500);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/70 backdrop-blur-sm animate-in fade-in duration-150">
      <div className="relative w-full max-w-md bg-slate-900 border border-slate-700/80 rounded-3xl p-6 shadow-2xl text-slate-100">
        <button
          onClick={onClose}
          className="absolute top-5 right-5 text-slate-400 hover:text-white p-1 rounded-full hover:bg-slate-800 transition"
        >
          <X className="w-5 h-5" />
        </button>

        {!isAuthenticated ? (
          <div>
            <div className="flex items-center gap-3 mb-4">
              <div className="w-10 h-10 rounded-xl bg-amber-500/10 border border-amber-500/20 text-amber-400 flex items-center justify-center">
                <Shield className="w-5 h-5" />
              </div>
              <div>
                <h3 className="font-bold text-lg">Admin Voucher Generator</h3>
                <p className="text-xs text-slate-400">Khusus pemilik (Bayu Nurcahyono)</p>
              </div>
            </div>

            <form onSubmit={handleLogin} className="space-y-4 mt-6">
              <div>
                <label className="block text-xs font-semibold text-slate-300 mb-1.5">
                  Masukkan PIN Rahasia Admin
                </label>
                <div className="relative">
                  <Key className="w-4 h-4 text-slate-500 absolute left-3.5 top-1/2 -translate-y-1/2" />
                  <input
                    type="password"
                    value={pin}
                    onChange={(e) => {
                      setPin(e.target.value);
                      setPinError(false);
                    }}
                    placeholder="Masukkan PIN rahasia..."
                    className="w-full pl-10 pr-4 py-2.5 bg-slate-800/80 border border-slate-700 rounded-xl text-sm focus:outline-none focus:border-amber-500 placeholder:text-slate-500"
                    autoFocus
                  />
                </div>
                {pinError && <p className="text-xs text-rose-400 mt-1">PIN salah! Coba lagi.</p>}
              </div>

              <button
                type="submit"
                className="w-full py-2.5 rounded-xl bg-amber-500 hover:bg-amber-400 text-slate-950 font-bold text-sm transition shadow-lg shadow-amber-500/20"
              >
                Buka Panel Admin
              </button>
            </form>
          </div>
        ) : (
          <div className="space-y-6">
            <div className="flex items-center gap-2">
              <div className="w-8 h-8 rounded-lg bg-emerald-500/20 text-emerald-400 flex items-center justify-center">
                <Sparkles className="w-4 h-4" />
              </div>
              <div>
                <h3 className="font-bold text-base">Buat Kode Voucher Pembeli</h3>
                <p className="text-xs text-slate-400">Generate kode voucher setelah pembeli transfer</p>
              </div>
            </div>

            {/* Tombol Pilihan Denominasi */}
            <div>
              <label className="block text-xs font-semibold text-slate-400 mb-2">PILIH DENOMINASI TOKEN:</label>
              <div className="grid grid-cols-3 gap-2">
                {[20, 50, 100].map((t) => (
                  <button
                    key={t}
                    type="button"
                    onClick={() => {
                      setSelectedTokens(t);
                      handleGenerate(t);
                    }}
                    className={`py-2.5 px-3 rounded-xl border text-sm font-bold flex flex-col items-center gap-0.5 transition ${
                      selectedTokens === t
                        ? 'bg-amber-500/20 border-amber-500 text-amber-300 shadow-md shadow-amber-500/10'
                        : 'bg-slate-800 border-slate-700 hover:border-slate-600 text-slate-300'
                    }`}
                  >
                    <span>{t} Token</span>
                    <span className="text-[10px] font-normal text-slate-400">
                      Rp {(t * 1000).toLocaleString('id-ID')}
                    </span>
                  </button>
                ))}
              </div>
            </div>

            {/* Hasil Kode */}
            {generatedCode ? (
              <div className="p-4 rounded-2xl bg-slate-950 border border-amber-500/30 space-y-2">
                <span className="text-[11px] font-semibold text-amber-400 uppercase tracking-wider">
                  Kode Voucher Siap Kirim
                </span>
                <div className="flex items-center justify-between gap-2">
                  <code className="text-base font-mono font-bold text-white tracking-wider">
                    {generatedCode}
                  </code>
                  <button
                    onClick={() => copyToClipboard(generatedCode)}
                    className="p-2 rounded-lg bg-amber-500 text-slate-950 hover:bg-amber-400 transition font-medium flex items-center gap-1.5 text-xs shrink-0"
                  >
                    {copied ? <Check className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
                    {copied ? 'Tersalin!' : 'Salin'}
                  </button>
                </div>
                <p className="text-[11px] text-slate-400">
                  Kirimkan kode ini ke pembeli di WhatsApp untuk di-redeem di aplikasi.
                </p>
              </div>
            ) : (
              <button
                onClick={() => handleGenerate(selectedTokens)}
                className="w-full py-2.5 rounded-xl bg-amber-500 hover:bg-amber-400 text-slate-950 font-bold text-sm transition"
              >
                Generate Voucher {selectedTokens} Token
              </button>
            )}

            {/* Shortcut Tambah Langsung ke Browser Ini */}
            <div className="pt-4 border-t border-slate-800">
              <div className="flex items-center justify-between text-xs mb-2">
                <span className="text-slate-400 font-semibold">TAMBAH TOKEN KE AKUN SENDIRI:</span>
                {directAddSuccess && <span className="text-emerald-400 font-medium">Berhasil ditambahkan!</span>}
              </div>
              <div className="flex gap-2">
                <select
                  value={directAddAmount}
                  onChange={(e) => setDirectAddAmount(Number(e.target.value))}
                  className="bg-slate-800 border border-slate-700 text-slate-200 text-xs rounded-xl px-3 py-2 focus:outline-none"
                >
                  <option value={10}>+10 Token</option>
                  <option value={50}>+50 Token</option>
                  <option value={100}>+100 Token</option>
                  <option value={500}>+500 Token</option>
                </select>
                <button
                  type="button"
                  onClick={handleDirectAdd}
                  className="flex-1 py-2 px-3 rounded-xl bg-slate-700 hover:bg-slate-600 text-white font-medium text-xs flex items-center justify-center gap-1.5 transition"
                >
                  <PlusCircle className="w-4 h-4 text-emerald-400" />
                  Isi ke Akun Saya
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
