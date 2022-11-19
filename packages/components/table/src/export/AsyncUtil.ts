export class AsyncUtil {
  public static execWithListDataByAsync<T>(
    lstData: Array<T>,
    consumer: (data: T) => void
  ): Promise<void> {
    return new Promise((resolve) => {
      const lstPromise: Array<Promise<void>> = []
      for (const data of lstData) {
        lstPromise.push(
          new Promise((resolve) => {
            consumer(data)
            resolve()
          })
        )
      }
      Promise.all(lstPromise).then(() => {
        resolve()
      })
    })
  }
}
