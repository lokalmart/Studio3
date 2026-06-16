export default function AssistantPage() {
  return (
    <main className="min-h-screen px-4 py-6">
      <div className="card mx-auto max-w-3xl p-6">
        <h1 className="text-3xl font-black">Asisten Lokalmart</h1>
        <p className="mt-3 text-orange-100/70">
          Di v8, Asisten tidak lagi menjadi fitur utama di home. Alur utama Studio2 adalah Import dan Export.
          Gunakan export project/model untuk membuat XLSX konteks, lalu berikan file itu ke ChatGPT untuk dianalisis.
        </p>
        <a href="/" className="btn btn-primary mt-5">Kembali ke Studio2</a>
      </div>
    </main>
  );
}
