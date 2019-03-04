package main

import (
	"fmt"
	"sort"
	"strconv"
)

func main() {
	word := "AAA123123"
	wi := 0
	for _, w := range word {
		if (('a' <= w && w <= 'z') || ('A' <= w && w <= 'Z')) {
			wi++
			continue
		}
		break
	}
	num, _ := strconv.Atoi(word[wi:len(word)])
	print(word[0:wi], " ", num);

	var us [] string
	const N = 6
	for i := 0; i < N; i++ {
		us = append(us, strconv.Itoa(6-i))
	}
	sort.Strings(us)
	fmt.Println(us)
}
