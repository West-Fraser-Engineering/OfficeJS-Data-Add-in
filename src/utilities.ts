export async function Delay(ms?: number) {
    return new Promise<void>(res => {
        setTimeout(res, ms);
    });
}