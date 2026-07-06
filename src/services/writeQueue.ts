/** 按工作簿路径保存的写入队列。 */
const workbookQueues = new Map<string, Promise<unknown>>();

/** 对同一个工作簿路径串行执行写入操作。 */
export async function withWorkbookWrite<T>(workbookPath: string, operation: () => Promise<T>): Promise<T> {
  const previous = workbookQueues.get(workbookPath) ?? Promise.resolve();
  const current = previous.then(operation, operation);
  workbookQueues.set(
    workbookPath,
    current.finally(() => {
      if (workbookQueues.get(workbookPath) === current) {
        workbookQueues.delete(workbookPath);
      }
    }),
  );
  return current;
}
