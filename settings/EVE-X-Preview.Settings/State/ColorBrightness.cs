namespace EveX.Settings.State;

// Preserve RGB ratios while changing HSV value, including a round trip through black.
public sealed class ColorBrightness
{
    private double red = 1, green = 1, blue = 1;
    public double Percent { get; private set; }
    public void SetRgb(byte r, byte g, byte b)
    {
        var maximum = Math.Max(r, Math.Max(g, b));
        Percent = maximum * 100d / 255;
        if (maximum == 0) return;
        red = r / (double)maximum; green = g / (double)maximum; blue = b / (double)maximum;
    }
    public (byte r, byte g, byte b) WithValue(double percent)
    {
        Percent = Math.Clamp(percent, 0, 100);
        var maximum = Percent * 255 / 100;
        return ((byte)Math.Round(red * maximum), (byte)Math.Round(green * maximum), (byte)Math.Round(blue * maximum));
    }
}
