using System;
using System.Threading;
using Microsoft.Extensions.Hosting;

namespace FileOrganizer;

internal sealed class HeartbeatService : IDisposable
{
    private readonly IHostApplicationLifetime _lifetime;
    private readonly Timer _timer;
    private long _lastBeatTicks;

    public HeartbeatService(IHostApplicationLifetime lifetime)
    {
        _lifetime = lifetime;
        Beat();
        _timer = new Timer(CheckIdleWindow, null, TimeSpan.FromSeconds(15), TimeSpan.FromSeconds(15));
    }

    public void Beat()
    {
        Interlocked.Exchange(ref _lastBeatTicks, DateTimeOffset.UtcNow.UtcTicks);
    }

    public void Dispose()
    {
        _timer.Dispose();
    }

    private void CheckIdleWindow(object? _)
    {
        var lastBeat = new DateTimeOffset(Interlocked.Read(ref _lastBeatTicks), TimeSpan.Zero);
        if (DateTimeOffset.UtcNow - lastBeat > TimeSpan.FromHours(12))
        {
            _lifetime.StopApplication();
        }
    }
}

