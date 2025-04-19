/**
 * Utility functions for color manipulation based on Fluent UI button color behavior
 * 
 * When used with the default Fluent UI primary blue (#1267b4),
 * these calculations produce:
 * - Hover: #18599b (darker blue with reduced saturation)
 * - Click: #193253 (much darker blue with further reduced saturation)
 */
export const ColorUtils = {
    /**
     * Normalizes a hex color string to ensure it has a hashtag prefix
     * @param color - The hex color string with or without # prefix
     * @returns The color with a # prefix
     */
    normalizeColor: (color: string): string => {
        return color.startsWith('#') ? color : `#${color}`;
    },

    /**
     * Determines whether a color is dark enough to require white text for contrast
     * @param hexColor - The background color as a hex string
     * @returns true if the color should have white text, false for black text
     */
    shouldUseWhiteText: (hexColor: string): boolean => {
        // Normalize the color
        const normalizedColor = ColorUtils.normalizeColor(hexColor);
        const rgb = ColorUtils.hexToRgb(normalizedColor);

        // Calculate relative luminance using the sRGB formula (WCAG 2.0)
        // https://www.w3.org/WAI/GL/wiki/Relative_luminance
        const r = rgb.r / 255;
        const g = rgb.g / 255;
        const b = rgb.b / 255;

        const R = r <= 0.03928 ? r / 12.92 : Math.pow((r + 0.055) / 1.055, 2.4);
        const G = g <= 0.03928 ? g / 12.92 : Math.pow((g + 0.055) / 1.055, 2.4);
        const B = b <= 0.03928 ? b / 12.92 : Math.pow((b + 0.055) / 1.055, 2.4);

        // Calculate luminance
        const luminance = 0.2126 * R + 0.7152 * G + 0.0722 * B;

        // Use white text for darker colors (lower luminance)
        return luminance < 0.5;
    },
    /**
     * Generates hover color based on base color
     * Using a formula calibrated to match the Fluent UI default hover effect
     */    getHoverColor: (baseColor: string): string => {
        // Normalize the color
        const normalizedColor = ColorUtils.normalizeColor(baseColor);

        // Convert the base color to HSL for better control over the transformation
        const rgb = ColorUtils.hexToRgb(normalizedColor);
        const hsl = ColorUtils.rgbToHsl(rgb.r, rgb.g, rgb.b);

        // Hover effect: slightly increase saturation, decrease lightness
        // Parameters calibrated to match #1267b4 → #18599b transformation
        hsl.h += 0.01;  // Slight hue shift
        hsl.s *= 0.92;  // Reduce saturation by 8%
        hsl.l *= 0.86;  // Reduce lightness by 14%

        // Convert back to RGB and then to hex
        const newRgb = ColorUtils.hslToRgb(hsl.h, hsl.s, hsl.l);
        return ColorUtils.rgbToHex(newRgb.r, newRgb.g, newRgb.b);
    },

    /**
     * Generates pressed color based on base color
     * Using a formula calibrated to match the Fluent UI default pressed effect
     */    getPressedColor: (baseColor: string): string => {
        // Normalize the color
        const normalizedColor = ColorUtils.normalizeColor(baseColor);

        // Convert the base color to HSL for better control over the transformation
        const rgb = ColorUtils.hexToRgb(normalizedColor);
        const hsl = ColorUtils.rgbToHsl(rgb.r, rgb.g, rgb.b);

        // Pressed effect: decrease saturation, significantly decrease lightness
        // Parameters calibrated to match #1267b4 → #193253 transformation
        hsl.h += 0.02;   // Slightly more hue shift than hover
        hsl.s *= 0.75;   // Reduce saturation by 25%
        hsl.l *= 0.56;   // Reduce lightness by 44%

        // Convert back to RGB and then to hex
        const newRgb = ColorUtils.hslToRgb(hsl.h, hsl.s, hsl.l);
        return ColorUtils.rgbToHex(newRgb.r, newRgb.g, newRgb.b);
    },

    /**
     * Convert hex color to RGB
     */
    hexToRgb: (hex: string): { r: number; g: number; b: number } => {
        // Handle colors with or without # prefix
        const cleanHex = hex.startsWith('#') ? hex.slice(1) : hex;

        const r = parseInt(cleanHex.substring(0, 2), 16);
        const g = parseInt(cleanHex.substring(2, 4), 16);
        const b = parseInt(cleanHex.substring(4, 6), 16);

        return { r, g, b };
    },

    /**
     * Convert RGB to hex color string
     */
    rgbToHex: (r: number, g: number, b: number): string => {
        const rr = Math.round(Math.max(0, Math.min(255, r))).toString(16).padStart(2, '0');
        const gg = Math.round(Math.max(0, Math.min(255, g))).toString(16).padStart(2, '0');
        const bb = Math.round(Math.max(0, Math.min(255, b))).toString(16).padStart(2, '0');

        return `#${rr}${gg}${bb}`;
    },

    /**
     * Convert RGB to HSL
     * Returns object with h, s, l properties each in range [0, 1]
     */
    rgbToHsl: (r: number, g: number, b: number): { h: number; s: number; l: number } => {
        r /= 255;
        g /= 255;
        b /= 255;

        const max = Math.max(r, g, b);
        const min = Math.min(r, g, b);
        let h = 0;
        let s = 0;
        const l = (max + min) / 2;

        if (max !== min) {
            const d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);

            switch (max) {
                case r: h = (g - b) / d + (g < b ? 6 : 0); break;
                case g: h = (b - r) / d + 2; break;
                case b: h = (r - g) / d + 4; break;
            }

            h /= 6;
        }

        return { h, s, l };
    },

    /**
     * Convert HSL to RGB
     * Expects h, s, l each in range [0, 1]
     * Returns object with r, g, b properties each in range [0, 255]
     */
    hslToRgb: (h: number, s: number, l: number): { r: number; g: number; b: number } => {
        let r, g, b;

        if (s === 0) {
            r = g = b = l; // achromatic
        } else {
            const hue2rgb = (p: number, q: number, t: number): number => {
                if (t < 0) t += 1;
                if (t > 1) t -= 1;
                if (t < 1 / 6) return p + (q - p) * 6 * t;
                if (t < 1 / 2) return q;
                if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
                return p;
            };

            const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
            const p = 2 * l - q;

            r = hue2rgb(p, q, h + 1 / 3);
            g = hue2rgb(p, q, h);
            b = hue2rgb(p, q, h - 1 / 3);
        }

        return {
            r: r * 255,
            g: g * 255,
            b: b * 255
        };
    }
};
