#include <windows.h>
#include <stdio.h>
#include <string.h>

int main() {
    STARTUPINFO si;
    PROCESS_INFORMATION pi;
    char programPath[MAX_PATH];
    char *programName = "pdf_function.exe";

    // ���݂̃v���O�����̃t���p�X���擾
    DWORD length = GetModuleFileName(NULL, programPath, sizeof(programPath));
    if (length == 0) {
        printf("GetModuleFileName failed (%d).\n", GetLastError());
        return -1;
    }

    // �t�@�C����������T���Ēu��������
    char *lastSlash = strrchr(programPath, '\\');
    if (lastSlash) {
        strcpy(lastSlash + 1, programName);  // �t�@�C����������u������
    } else {
        printf("Could not find slash in path.\n");
        return -1;
    }

    ZeroMemory(&si, sizeof(si));
    si.cb = sizeof(si);
    ZeroMemory(&pi, sizeof(pi));

    if (!CreateProcess(NULL,
                        programPath,  // �u���������v���O���������g�p
                        NULL,
                        NULL,
                        FALSE,
                        0,  // �����̃R���\�[�����g�p
                        NULL,
                        NULL,
                        &si,
                        &pi)
    ) {
        printf("CreateProcess failed (%d).\n", GetLastError());
        return -1;
    }

    // �e�v���Z�X�͂����ɏI������
    return 0;
}
