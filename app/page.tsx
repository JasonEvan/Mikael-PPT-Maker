import Image from "next/image";

// ponytail: styled with standard Bootstrap classes
export default function Home() {
  return (
    <div className="container min-vh-100 d-flex flex-column align-items-center justify-content-center py-5">
      <main className="d-flex flex-column align-items-center align-items-md-start gap-4 p-5 bg-light rounded-4 shadow-sm">
        <Image
          src="/next.svg"
          alt="Next.js logo"
          width={100}
          height={20}
          priority
        />
        <div className="text-center text-md-start">
          <h1 className="h3 fw-bold mb-3">
            To get started, edit the{" "}
            <code className="bg-body-secondary px-2 py-1 rounded">
              app/page.tsx
            </code>{" "}
            file.
          </h1>
          <p className="lead text-secondary">
            Looking for a starting point or more instructions? Head over to{" "}
            <a
              href="https://vercel.com/templates?framework=next.js"
              className="text-decoration-none fw-semibold"
            >
              Templates
            </a>{" "}
            or the{" "}
            <a
              href="https://nextjs.org/learn"
              className="text-decoration-none fw-semibold"
            >
              Learning
            </a>{" "}
            center.
          </p>
        </div>
        <div className="d-flex flex-column flex-sm-row gap-3 w-100">
          <a
            className="btn btn-dark rounded-pill px-4 py-2 d-flex align-items-center justify-center gap-2"
            href="https://vercel.com/new"
            target="_blank"
            rel="noopener noreferrer"
          >
            <Image
              src="/vercel.svg"
              alt="Vercel logomark"
              width={16}
              height={14}
            />
            Deploy Now
          </a>
          <a
            className="btn btn-outline-secondary rounded-pill px-4 py-2"
            href="https://nextjs.org/docs"
            target="_blank"
            rel="noopener noreferrer"
          >
            Documentation
          </a>
        </div>
      </main>
    </div>
  );
}
