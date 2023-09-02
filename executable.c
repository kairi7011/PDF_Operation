#include <windows.h>
#include <stdio.h>
#include <string.h>

int main() {
    STARTUPINFO si;
    PROCESS_INFORMATION pi;
    char programPath[MAX_PATH];
    char *programName = "pdf_function.exe";

    // 現在のプログラムのフルパスを取得
    DWORD length = GetModuleFileName(NULL, programPath, sizeof(programPath));
    if (length == 0) {
        printf("GetModuleFileName failed (%d).\n", GetLastError());
        return -1;
    }

    // ファイル名部分を探して置き換える
    char *lastSlash = strrchr(programPath, '\\');
    if (lastSlash) {
        strcpy(lastSlash + 1, programName);  // ファイル名部分を置き換え
    } else {
        printf("Could not find slash in path.\n");
        return -1;
    }

    ZeroMemory(&si, sizeof(si));
    si.cb = sizeof(si);
    ZeroMemory(&pi, sizeof(pi));

    if (!CreateProcess(NULL,
                        programPath,  // 置き換えたプログラム名を使用
                        NULL,
                        NULL,
                        FALSE,
                        0,  // 既存のコンソールを使用
                        NULL,
                        NULL,
                        &si,
                        &pi)
    ) {
        printf("CreateProcess failed (%d).\n", GetLastError());
        return -1;
    }

    // 親プロセスはすぐに終了する
    return 0;
}
