interface WMF2PNGInstance {
  getPNG(file: File | string): Promise<string>
  getBase64(file: File): Promise<string>
  transformWMF(base64: string): string
}

declare const WMF2PNG: WMF2PNGInstance

export default WMF2PNG
