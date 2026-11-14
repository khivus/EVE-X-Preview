namespace EveX.Settings.State;

// Presentation colors only; the canonical theme palettes remain owned by AHK.
public static class SelectionContrast
{
    public static double Ratio(uint first, uint second)
    {
        static double Luminance(uint color)
        {
            static double Linear(uint value) { var c = value / 255d; return c <= .04045 ? c / 12.92 : Math.Pow((c + .055) / 1.055, 2.4); }
            return .2126 * Linear((color >> 16) & 255) + .7152 * Linear((color >> 8) & 255) + .0722 * Linear(color & 255);
        }
        var a = Luminance(first); var b = Luminance(second);
        return (Math.Max(a, b) + .05) / (Math.Min(a, b) + .05);
    }

    public static uint Text(uint background, uint preferred) => Ratio(background, preferred) >= 4.5
        ? preferred : Ratio(background, 0) > Ratio(background, 0xFFFFFF) ? 0u : 0xFFFFFFu;

    public static uint Outline(uint accent, uint background, uint input)
    {
        double Visibility(uint color) => Math.Min(Ratio(color, background), Ratio(color, input));
        return Visibility(accent) >= 3 ? accent : Visibility(0) > Visibility(0xFFFFFF) ? 0u : 0xFFFFFFu;
    }

    public static uint Fill(uint input, uint accent)
    {
        uint Channel(int shift) => (uint)Math.Round(((input >> shift) & 255) * .65 + ((accent >> shift) & 255) * .35);
        return (Channel(16) << 16) | (Channel(8) << 8) | Channel(0);
    }
}
