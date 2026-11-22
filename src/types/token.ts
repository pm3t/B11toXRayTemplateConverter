export interface BankAccount {
  bank: string;
  number: string;
  holder: string;
  logoColor?: string;
  iconName?: string;
}

export interface TokenPackage {
  id: string;
  tokens: number;
  price: number;
  formattedPrice: string;
  isPopular?: boolean;
  description: string;
}

export interface TokenTransaction {
  id: string;
  timestamp: string;
  type: 'bonus' | 'redeem' | 'conversion' | 'admin_adjustment';
  amount: number; // positive for addition, negative for consumption
  balanceAfter: number;
  description: string;
  referenceCode?: string;
}

export interface UserWallet {
  userId: string;
  balance: number;
  createdAt: string;
  totalConversions: number;
  transactions: TokenTransaction[];
  hasReceivedFreeBonus: boolean;
}
