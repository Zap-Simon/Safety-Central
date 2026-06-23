import { Pool, neonConfig } from '@neondatabase/serverless';
import { drizzle } from 'drizzle-orm/neon-serverless';
import ws from "ws";
import * as schema from "@shared/schema";

neonConfig.webSocketConstructor = ws;

if (!process.env.DATABASE_URL) {
  throw new Error(
    "DATABASE_URL must be set. Did you forget to provision a database?",
  );
}

export const pool = new Pool({ connectionString: process.env.DATABASE_URL });

// Neon's serverless Pool emits 'error' on idle clients whose underlying
// WebSocket drops (e.g. after a failed query or a network blip). If left
// unhandled this surfaces as an uncaught exception and can leave the pool in a
// bad state, stalling every later DB-backed request. Logging it lets the pool
// discard the broken client and hand out a fresh one on the next acquire.
pool.on('error', (err) => {
  console.error('Database pool error (idle client):', err);
});

export const db = drizzle({ client: pool, schema });
