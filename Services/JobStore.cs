using System.Collections.Concurrent;

namespace DataDownloaderContPaqi.Web.Services;

// Holds completed job results in memory until the browser downloads them.
// Entries are removed on first retrieval to avoid unbounded growth.
public class JobStore
{
    private readonly ConcurrentDictionary<Guid, (byte[] Data, string FileName, string ContentType)> _jobs = new();

    public Guid Store(byte[] data, string fileName, string contentType)
    {
        var id = Guid.NewGuid();
        _jobs[id] = (data, fileName, contentType);
        return id;
    }

    public bool TryGet(Guid id, out byte[] data, out string fileName, out string contentType)
    {
        if (_jobs.TryGetValue(id, out var entry))
        {
            data = entry.Data;
            fileName = entry.FileName;
            contentType = entry.ContentType;
            return true;
        }
        data = [];
        fileName = string.Empty;
        contentType = string.Empty;
        return false;
    }

    public void Remove(Guid id) => _jobs.TryRemove(id, out _);
}
